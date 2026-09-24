// Builds the Pharma Sector object -- the exact shape App.tsx already
// expects (see src/data/sectors.ts) -- from live Neo4j queries instead
// of parsing the static pharma.xlsx. Nothing downstream changes: the
// Overview, Process maps, and Data ontology tabs all keep working
// exactly as before, just reading live data now.
'use strict';

const neo4j = require('neo4j-driver');

const NEO4J_URI = process.env.NEO4J_URI || 'bolt://localhost:7687';
const NEO4J_USER = process.env.NEO4J_USER || 'neo4j';
const NEO4J_PASSWORD = process.env.NEO4J_PASSWORD || 'Sunflower@123';

function slug(value) {
  return String(value)
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
}

function capabilityKey(functionName, capabilityName) {
  return `${functionName}\u0000${capabilityName}`;
}

function capabilityId(sectorId, functionName, capabilityName) {
  return `${sectorId}-${slug(functionName)}-${slug(capabilityName)}`;
}

async function runQuery(session, cypher, params = {}) {
  const result = await session.run(cypher, params);
  return result.records.map((r) => r.toObject());
}

async function getPharmaSector() {
  // disableLosslessIntegers: without this, whole numbers stored in Neo4j
  // (like l1.order) come back as Integer *objects*, not plain numbers --
  // which is exactly what was rendering as "[object Object]" in the UI.
  const driver = neo4j.driver(
    NEO4J_URI,
    neo4j.auth.basic(NEO4J_USER, NEO4J_PASSWORD),
    { disableLosslessIntegers: true }
  );
  const session = driver.session();

  try {
    // --- L1 stages ---
    const l1Rows = await runQuery(
      session,
      'MATCH (l1:L1) RETURN l1.id AS id, l1.order AS order, l1.name AS name ORDER BY l1.order'
    );
    const stages = l1Rows.map((r) => ({
      id: r.id,
      order: r.order,
      name: r.name,
      description: r.name, // L1 description wasn't loaded into Neo4j; name is a safe fallback
    }));

    // --- Function -> Capability -> L3 (via USES_CAPABILITY) -> L2 -> L1 ---
    // This reconstructs the functions[].capabilities[].processes[] tree.
    const taxonomyRows = await runQuery(
      session,
      `MATCH (f:Function)-[:HAS_CAPABILITY]->(c:Capability)
       OPTIONAL MATCH (l3:L3)-[:USES_CAPABILITY]->(c) WHERE l3.primaryFunction = f.name
       OPTIONAL MATCH (l2:L2)-[:CONTAINS]->(l3)
       OPTIONAL MATCH (l1:L1)-[:CONTAINS]->(l2)
       RETURN f.name AS functionName, c.name AS capabilityName,
              l3.id AS l3Id, l3.name AS l3Name,
              l2.id AS l2Id, l2.name AS l2Name,
              l1.id AS l1Id, l1.name AS l1Name`
    );

    // --- Entities (with their process context resolved) ---
    const entityRows = await runQuery(
      session,
      `MATCH (e:Entity)
       OPTIONAL MATCH (l3:L3 {id: e.primaryL3})
       OPTIONAL MATCH (l2:L2)-[:CONTAINS]->(l3)
       OPTIONAL MATCH (l1:L1)-[:CONTAINS]->(l2)
       RETURN e.id AS rawId, e.description AS description, e.system AS system,
              e.fields AS fields, e.primaryFunction AS primaryFunction,
              e.primaryCapability AS primaryCapability,
              e.primaryL3 AS l3Id, l3.name AS l3Name,
              l2.id AS l2Id, l2.name AS l2Name,
              l1.id AS l1Id, l1.name AS l1Name`
    );

    // --- Entity-to-entity relationships (both directions) ---
    const relRows = await runQuery(
      session,
      'MATCH (a:Entity)-[r]->(b:Entity) RETURN a.id AS fromId, type(r) AS relationship, b.id AS toId'
    );

    const relationshipsByEntity = new Map();
    const entityRelationships = [];
    relRows.forEach(({ fromId, relationship, toId }) => {
      entityRelationships.push({ from: fromId, relationship, to: toId });
      const phrase = relationship.replace(/_/g, ' ').toLowerCase();
      const outgoing = relationshipsByEntity.get(fromId) ?? [];
      outgoing.push(`${phrase} ${toId}`);
      relationshipsByEntity.set(fromId, outgoing);
      const incoming = relationshipsByEntity.get(toId) ?? [];
      incoming.push(`${fromId} ${phrase}`);
      relationshipsByEntity.set(toId, incoming);
    });

    // --- Build the processes, grouped under their capability ---
    const processesByCapability = new Map();
    const seenL3 = new Set();
    taxonomyRows.forEach((row) => {
      if (!row.l3Id || seenL3.has(row.l3Id + row.functionName)) return;
      seenL3.add(row.l3Id + row.functionName);
      const key = capabilityKey(row.functionName, row.capabilityName);
      const dataObjectIds = entityRows
        .filter((e) => e.l3Id === row.l3Id)
        .map((e) => slug(e.rawId));
      const process = {
        id: row.l3Id,
        name: row.l3Name,
        description: '',
        owner: row.functionName,
        stageId: row.l1Id,
        stageName: row.l1Name,
        parentProcessId: row.l2Id,
        parentProcessName: row.l2Name,
        contributingFunctions: [],
        contributingCapabilities: [],
        dataObjectIds,
        inputs: [],
        outputs: [],
        steps: [row.l3Name],
      };
      const list = processesByCapability.get(key) ?? [];
      list.push(process);
      processesByCapability.set(key, list);
    });

    // --- Build functions[].capabilities[] ---
    const functionOrder = [];
    const groupedTaxonomy = new Map();
    taxonomyRows.forEach(({ functionName, capabilityName }) => {
      if (!functionName || !capabilityName) return;
      if (!functionOrder.includes(functionName)) functionOrder.push(functionName);
      const caps = groupedTaxonomy.get(functionName) ?? [];
      if (!caps.includes(capabilityName)) caps.push(capabilityName);
      groupedTaxonomy.set(functionName, caps);
    });

    const functions = functionOrder
      .map((functionName) => ({
        id: `pharma-${slug(functionName)}`,
        name: functionName,
        description: `Processes and governed data owned by ${functionName}.`,
        owner: functionName,
        capabilities: (groupedTaxonomy.get(functionName) ?? [])
          .map((capabilityName) => {
            const processList = processesByCapability.get(capabilityKey(functionName, capabilityName)) ?? [];
            const systems = new Set(
              entityRows
                .filter((e) => e.primaryFunction === functionName && e.primaryCapability === capabilityName)
                .map((e) => e.system)
                .filter(Boolean)
            );
            return {
              id: capabilityId('pharma', functionName, capabilityName),
              name: capabilityName,
              description: processList[0]?.parentProcessName
                ? `Includes ${processList[0].parentProcessName} and related sub-steps.`
                : `Operational processes for ${capabilityName}.`,
              systems: Array.from(systems),
              processes: processList,
            };
          })
          .filter((c) => c.processes.length > 0),
      }))
      .filter((f) => f.capabilities.length > 0);

    // --- Build entities[] in the exact DataEntity shape ---
    const entities = entityRows.map((row) => ({
      id: slug(row.rawId),
      name: row.rawId, // raw underscored code -- the UI formats this for display
      description: row.description ?? '',
      system: row.system ?? '',
      capabilityId: capabilityId('pharma', row.primaryFunction ?? '', row.primaryCapability ?? ''),
      attributes: row.fields ?? [],
      relationships: relationshipsByEntity.get(row.rawId) ?? [],
      primaryFunction: row.primaryFunction ?? undefined,
      primaryCapability: row.primaryCapability ?? undefined,
      l1Id: row.l1Id ?? undefined,
      l1Name: row.l1Name ?? undefined,
      l2Id: row.l2Id ?? undefined,
      l2Name: row.l2Name ?? undefined,
      l3Id: row.l3Id ?? undefined,
      l3Name: row.l3Name ?? undefined,
    }));

    return { stages, functions, entities, entityRelationships };
  } finally {
    await session.close();
    await driver.close();
  }
}

module.exports = { getPharmaSector };
