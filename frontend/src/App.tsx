import { useEffect, useRef, useState } from "react";
import {
  ArrowRight,
  BookOpen,
  Boxes,
  Check,
  ChevronDown,
  Database,
  Download,
  CupSoda,
  Drill,
  GitBranch,
  LoaderCircle,
  Network,
  Pill,
  Search,
  Upload,
  X,
} from "lucide-react";
import "./App.css";
import {
  sectorCatalog as defaultCatalog,
  uniqueCapabilityCount,
  type DataEntity,
  type Sector,
} from "./data/sectors";

const STORAGE_KEY = "industry-navigator-excel-catalog-v2";
type View = "overview" | "processes" | "ontology";
const sectorIcons = {
  pharma: Pill,
  "food-beverage": CupSoda,
  "oil-gas": Drill,
};

function loadCatalog() {
  return defaultCatalog;
}

async function readSectorWorkbook(identity: Sector) {
  // Pharma is now backed by Neo4j (live graph) instead of the static
  // workbook -- the API returns the exact same Sector shape, so nothing
  // downstream (Overview, Process maps, Data ontology) needs to change.
  if (identity.id === "pharma") {
    const response = await fetch(`/api/sector/pharma?updated=${Date.now()}`, {
      cache: "no-store",
    });
    if (!response.ok)
      throw new Error(`Unable to load ${identity.name} from Neo4j (${response.status}).`);
    const data = await response.json();
    return { ...identity, ...data } as Sector;
  }

  const { parseSectorWorkbook } = await import("./data/excelCatalog");
  const workbookUrl = `${import.meta.env.BASE_URL}data/${identity.id}.xlsx?updated=${Date.now()}`;
  const response = await fetch(workbookUrl, { cache: "no-store" });
  if (!response.ok)
    throw new Error(
      `Unable to load ${identity.name} workbook (${response.status}).`,
    );
  return parseSectorWorkbook(await response.arrayBuffer(), identity);
}

function App() {
  const [catalog, setCatalog] = useState<Sector[]>(loadCatalog);
  const [sectorId, setSectorId] = useState<string | null>(null);
  const [view, setView] = useState<View>("overview");
  const [, setCapabilityId] = useState("");
  const [expandedProcess, setExpandedProcess] = useState<string | null>(null);
  const [search, setSearch] = useState("");
  const [selectedEntity, setSelectedEntity] = useState<DataEntity | null>(null);
  const [loadingSector, setLoadingSector] = useState<string | null>(null);
  const [catalogLoading, setCatalogLoading] = useState(true);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const importRef = useRef<HTMLInputElement>(null);

  const sector = catalog.find((item) => item.id === sectorId);

  useEffect(() => {
    let cancelled = false;
    void Promise.all(
      defaultCatalog.map(async (identity) => {
        try {
          return await readSectorWorkbook(identity);
        } catch {
          return { ...identity, functions: [], entities: [] };
        }
      }),
    ).then((loadedCatalog) => {
      if (!cancelled) {
        setCatalog(loadedCatalog);
        setCatalogLoading(false);
      }
    });
    return () => {
      cancelled = true;
    };
  }, []);

  function saveCatalog(nextCatalog: Sector[]) {
    setCatalog(nextCatalog);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(nextCatalog));
  }

  function activateSector(nextSector: Sector) {
    setSectorId(nextSector.id);
    setCapabilityId(nextSector.functions[0]?.capabilities[0]?.id ?? "");
    setExpandedProcess(null);
    setSelectedEntity(null);
    setSearch("");
    setView("overview");
  }

  async function chooseSector(nextId: Sector["id"]) {
    const identity = catalog.find((item) => item.id === nextId);
    if (!identity) return;
    setError("");
    setLoadingSector(nextId);
    try {
      const parsed = await readSectorWorkbook(identity);
      saveCatalog(catalog.map((item) => (item.id === nextId ? parsed : item)));
      activateSector(parsed);
    } catch (loadError) {
      setError(
        loadError instanceof Error
          ? loadError.message
          : "The sector workbook could not be loaded.",
      );
    } finally {
      setLoadingSector(null);
    }
  }

  async function importWorkbook(file: File) {
    if (!sector) return;
    setError("");
    setMessage("Reading workbook...");
    try {
      const { parseSectorWorkbook } = await import("./data/excelCatalog");
      const parsed = await parseSectorWorkbook(
        await file.arrayBuffer(),
        sector,
      );
      saveCatalog(
        catalog.map((item) => (item.id === sector.id ? parsed : item)),
      );
      activateSector(parsed);
      setMessage(`${file.name} imported`);
      window.setTimeout(() => setMessage(""), 2600);
    } catch (importError) {
      setMessage("");
      setError(
        importError instanceof Error
          ? importError.message
          : "The Excel workbook could not be imported.",
      );
    }
  }

  async function exportWorkbook() {
    if (!sector) return;
    setError("");
    setMessage("Preparing workbook...");
    try {
      const { writeSectorWorkbook } = await import("./data/excelCatalog");
      const bytes = await writeSectorWorkbook(sector);
      const content = bytes.buffer.slice(
        bytes.byteOffset,
        bytes.byteOffset + bytes.byteLength,
      ) as ArrayBuffer;
      const url = URL.createObjectURL(
        new Blob([content], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
      );
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `${sector.id}-process-data-ontology.xlsx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      setMessage("Excel workbook exported");
      window.setTimeout(() => setMessage(""), 2200);
    } catch (exportError) {
      setMessage("");
      setError(
        exportError instanceof Error
          ? exportError.message
          : "The workbook could not be exported.",
      );
    }
  }

  if (!sector) {
    return (
      <SectorGate
        catalog={catalog}
        catalogLoading={catalogLoading}
        loadingSector={loadingSector}
        error={error}
        onChoose={(id) => void chooseSector(id)}
      />
    );
  }

  const substepCount = sector.functions.reduce(
    (total, item) =>
      total +
      item.capabilities.reduce(
        (sum, capability) => sum + capability.processes.length,
        0,
      ),
    0,
  );
  const processCount = new Set(
    sector.functions.flatMap((businessFunction) =>
      businessFunction.capabilities.flatMap((capability) =>
        capability.processes.map(
          (process) => process.parentProcessId ?? process.id,
        ),
      ),
    ),
  ).size;
  const stageCount =
    new Set((sector.stages ?? []).map((stage) => stage.id)).size || 1;
  const dataObjectCount = sector.entities.length;
  const capabilityCount = uniqueCapabilityCount(sector);

  return (
    <div
      className="app-shell"
      style={{ "--sector": sector.accent } as React.CSSProperties}
    >
      <header className="topbar">
        <button
          className="brand"
          type="button"
          onClick={() => setSectorId(null)}
        >
          <span className="brand-mark">
            <Network size={20} />
          </span>
          <span>Process Atlas</span>
        </button>
        <div className="sector-switcher">
          <span className="sector-dot" />
          <strong>{sector.name}</strong>
          <button type="button" onClick={() => setSectorId(null)}>
            Change sector
          </button>
        </div>
        <div className="top-actions">
          {message && (
            <span className="saved-note">
              <Check size={14} /> {message}
            </span>
          )}
          <button
            className="text-action"
            type="button"
            onClick={() => importRef.current?.click()}
          >
            <Upload size={16} /> Import Excel
          </button>
          <button
            className="data-button"
            type="button"
            onClick={() => void exportWorkbook()}
          >
            <Download size={16} /> Export Excel
          </button>
        </div>
      </header>

      <div className="workspace">
        <main className="main-content">
          {error && (
            <div className="error-banner">
              <X size={16} />
              {error}
              <button type="button" onClick={() => setError("")}>
                <X size={15} />
              </button>
            </div>
          )}
          <section className="workspace-header">
            <div>
              <span className="overline">{sector.eyebrow}</span>
              <h1>{sector.name}</h1>
              <p>{sector.description}</p>
            </div>
            <dl className="header-stats">
              <div>
                <dt>Value Chain</dt>
                <dd>{stageCount}</dd>
              </div>
              <div>
                <dt>Work Streams</dt>
                <dd>{processCount}</dd>
              </div>
              <div>
                <dt>Activities</dt>
                <dd>{substepCount}</dd>
              </div>
              <div>
                <dt>Data Objects</dt>
                <dd>{dataObjectCount}</dd>
              </div>
              <div>
                <dt>Functions</dt>
                <dd>{sector.functions.length}</dd>
              </div>
              <div>
                <dt>Capabilities</dt>
                <dd>{capabilityCount}</dd>
              </div>
            </dl>
          </section>

          <div className="view-tabs" role="tablist">
            <button
              className={view === "overview" ? "active" : ""}
              type="button"
              onClick={() => setView("overview")}
            >
              <BookOpen size={16} /> Overview
            </button>
            <button
              className={view === "processes" ? "active" : ""}
              type="button"
              onClick={() => setView("processes")}
            >
              <GitBranch size={16} /> Process maps
            </button>
            <button
              className={view === "ontology" ? "active" : ""}
              type="button"
              onClick={() => setView("ontology")}
            >
              <Database size={16} /> Data ontology
            </button>
          </div>

          {view === "overview" && (
            <Overview
              sector={sector}
              processCount={processCount}
              substepCount={substepCount}
              capabilityCount={capabilityCount}
            />
          )}
          {view === "processes" && (
            <ProcessExplorer
              sector={sector}
              expandedProcess={expandedProcess}
              onProcessToggle={(id) =>
                setExpandedProcess(expandedProcess === id ? null : id)
              }
            />
          )}
          {view === "ontology" && (
            <OntologyExplorer
              sector={sector}
              search={search}
              selectedEntity={selectedEntity}
              onSearch={setSearch}
              onSelect={setSelectedEntity}
            />
          )}
        </main>
      </div>
      <input
        className="visually-hidden"
        ref={importRef}
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) void importWorkbook(file);
          event.target.value = "";
        }}
      />
    </div>
  );
}

type SectorGateProps = {
  catalog: Sector[];
  catalogLoading: boolean;
  loadingSector: string | null;
  error: string;
  onChoose: (id: Sector["id"]) => void;
};
function SectorGate({
  catalog,
  catalogLoading,
  loadingSector,
  error,
  onChoose,
}: SectorGateProps) {
  return (
    <main className="sector-gate">
      <header className="gate-nav">
        <div className="brand">
          <span className="brand-mark">
            <Network size={20} />
          </span>
          <span>Process Atlas</span>
        </div>
      </header>
      <section className="gate-intro">
        <span className="overline">Process and data intelligence</span>
        <h1>Choose an industry lens</h1>
        <p>Navigate every layer. Deploy the right agent</p>
        {error && <div className="gate-error">{error}</div>}
      </section>
      <section className="sector-grid" aria-label="Available sectors">
        {catalog.map((item, index) => {
          const Icon = sectorIcons[item.id];
          const count = uniqueCapabilityCount(item);
          const loading = loadingSector === item.id;
          return (
            <button
              disabled={catalogLoading || Boolean(loadingSector)}
              className="sector-card"
              style={
                {
                  "--card-accent": item.accent,
                  "--delay": `${index * 90}ms`,
                } as React.CSSProperties
              }
              type="button"
              key={item.id}
              onClick={() => onChoose(item.id)}
            >
              <span className="sector-card-top">
                {loading || catalogLoading ? (
                  <LoaderCircle className="spin" size={28} />
                ) : (
                  <Icon size={28} />
                )}
                <span>0{index + 1}</span>
              </span>
              <span className="sector-card-copy">
                <strong>{item.name}</strong>
                <small>
                  {catalogLoading
                    ? "Reading sector workbook..."
                    : loading
                      ? "Reading sector workbook..."
                      : item.description}
                </small>
              </span>
              <span className="sector-card-meta">
                <span>
                  {catalogLoading
                    ? "Loading functions"
                    : `${item.functions.length} functions`}
                </span>
                <span>
                  {catalogLoading
                    ? "Loading capabilities"
                    : `${count} capabilities`}
                </span>
                <ArrowRight size={19} />
              </span>
            </button>
          );
        })}
      </section>
    </main>
  );
}

function Overview({
  sector,
  processCount,
  substepCount,
  capabilityCount,
}: {
  sector: Sector;
  processCount: number;
  substepCount: number;
  capabilityCount: number;
}) {
  const businessNotes = [
    `${sector.stages?.length ?? 1} value-chain chapters show how work moves from strategic intent to business outcomes.`,
    `${processCount} work streams organize the major business outcomes and ownership across the operating model.`,
    `${substepCount} activities define the executable work teams perform, including the handoffs and controls that keep delivery moving.`,
    `${sector.entities.length} data objects connect business work to the records, evidence, and decisions teams need to operate.`,
    `${sector.functions.length} functions and ${capabilityCount} capabilities provide the business taxonomy for assigning ownership and finding the right work.`,
  ];
  return (
    <section className="overview-panel">
      <div className="taxonomy-flow">
        <span>Cross-cutting taxonomy</span>
        <div>
          <strong>{sector.functions.length} functions</strong>
          <small>Accountable owners</small>
        </div>
        <div>
          <strong>{capabilityCount} capabilities</strong>
          <small>Unique business taxonomy</small>
        </div>
      </div>
      <div className="model-flow">
        <div>
          <span>L1</span>
          <strong>{sector.stages?.length ?? 1} Value Chain</strong>
          <small>Value-chain chapters</small>
        </div>
        <ArrowRight size={18} />
        <div>
          <span>L2</span>
          <strong>{processCount} Work Streams</strong>
          <small>Operational work</small>
        </div>
        <ArrowRight size={18} />
        <div>
          <span>L3</span>
          <strong>{substepCount} Activities</strong>
          <small>Executable work</small>
        </div>
        <ArrowRight size={18} />
        <div>
          <span>Data Objects</span>
          <strong>{sector.entities.length} Data Objects</strong>
          <small>Governed records</small>
        </div>
      </div>
      <div className="overview-columns">
        <section>
          <span className="label">Business operating model</span>
          <h2>How the business connects</h2>
          {businessNotes.map((line) => (
            <p key={line}>{line}</p>
          ))}
        </section>
        <section>
          <span className="label">Value Chain</span>
          <h2>Value Chain</h2>
          <div className="stage-list">
            {sector.stages?.map((stage) => (
              <div key={stage.id}>
                <span>{String(stage.order).padStart(2, "0")}</span>
                <div>
                  <strong>{stage.name}</strong>
                  <small>{businessStageDescription(stage.name)}</small>
                </div>
              </div>
            ))}
          </div>
        </section>
      </div>
    </section>
  );
}

function businessStageDescription(stageName: string) {
  const name = stageName.toLowerCase();
  if (name.includes("discovery"))
    return "Turn scientific opportunities into promising candidates and clear investment choices.";
  if (name.includes("clinical"))
    return "Generate reliable evidence, protect participants, and prepare the product for approval.";
  if (name.includes("analytical") || name.includes("process"))
    return "Build the product, process, and analytical foundations required for scale.";
  if (name.includes("transfer"))
    return "Move a proven process into manufacturing with the knowledge and controls teams need.";
  if (name.includes("regulatory"))
    return "Secure market authorization and manage the commitments that keep products available.";
  if (name.includes("manufacturing") || name.includes("supply"))
    return "Make, release, and deliver products reliably to meet patient and market demand.";
  if (name.includes("quality"))
    return "Protect product quality, resolve issues, and improve performance throughout the lifecycle.";
  if (name.includes("pharmacovigilance") || name.includes("safety"))
    return "Monitor product safety, respond to signals, and protect patients after launch.";
  return `Advance business outcomes through ${stageName.toLowerCase()}.`;
}

type ProcessExplorerProps = {
  sector: Sector;
  expandedProcess: string | null;
  onProcessToggle: (id: string) => void;
};
function ProcessExplorer({
  sector,
  expandedProcess,
  onProcessToggle,
}: ProcessExplorerProps) {
  const nodes = sector.functions.flatMap((businessFunction) =>
    businessFunction.capabilities.flatMap((capability) =>
      capability.processes.map((process) => ({
        businessFunction,
        capability,
        process: {
          ...process,
          parentProcessName: process.parentProcessName ?? process.name,
          description: process.description,
        },
      })),
    ),
  );
  const stages = (
    sector.stages ?? [
      {
        id: "S1",
        order: 1,
        name: sector.name,
        description: sector.description,
      },
    ]
  ).map((stage) => ({
    ...stage,
    name: stage.name,
  }));
  const [functionFilter, setFunctionFilter] = useState("all");
  const [capabilityFilter, setCapabilityFilter] = useState("all");
  const [collapsedStages, setCollapsedStages] = useState<Set<string>>(
    () => new Set(stages.map((stage) => stage.id)),
  );
  const [collapsedStreams, setCollapsedStreams] = useState<Set<string>>(
    new Set(),
  );
  const [selectedDataObject, setSelectedDataObject] =
    useState<DataEntity | null>(null);
  const [selectedDataObjectLevel, setSelectedDataObjectLevel] = useState<
    "l1" | "l2" | "l3" | null
  >(null);
  const [selectedDataObjectNode, setSelectedDataObjectNode] = useState<
    string | null
  >(null);
  function toggleCollapsed(
    setter: React.Dispatch<React.SetStateAction<Set<string>>>,
    id: string,
  ) {
    setter((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }
  function toggleStage(stageId: string, streamIds: string[]) {
    const willExpand = collapsedStages.has(stageId);
    toggleCollapsed(setCollapsedStages, stageId);
    if (willExpand)
      setCollapsedStreams(
        (current) =>
          new Set([
            ...current,
            ...streamIds.map((streamId) => `${stageId}-${streamId}`),
          ]),
      );
  }
  function dataObjectsForNode(node: {
    capability: Sector["functions"][number]["capabilities"][number];
    process: Sector["functions"][number]["capabilities"][number]["processes"][number];
  }) {
    const ids = node.process.dataObjectIds ?? [];
    return ids.length
      ? sector.entities.filter((entity) => ids.includes(entity.id))
      : sector.entities.filter(
          (entity) => entity.capabilityId === node.capability.id,
        );
  }
  function downloadDataObjectsJson(
    level: "L1" | "L2" | "L3",
    nodeName: string,
    dataObjects: DataEntity[],
  ) {
    const names = new Set(dataObjects.map((entity) => entity.name));
    const connections = (sector.entityRelationships ?? []).filter(
      (relationship) =>
        names.has(relationship.from) && names.has(relationship.to),
    );
    const payload = {
      level,
      node: nodeName,
      dataObjects: dataObjects.map((entity) => ({
        name: entity.name,
        attributes: Object.fromEntries(
          entity.attributes.map((attribute) => [attribute, ""]),
        ),
        connections: connections.filter(
          (connection) =>
            connection.from === entity.name || connection.to === entity.name,
        ),
      })),
    };
    const url = URL.createObjectURL(
      new Blob([JSON.stringify(payload, null, 2)], {
        type: "application/json",
      }),
    );
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${sector.id}-${level.toLowerCase()}-${nodeName.toLowerCase().replace(/[^a-z0-9]+/g, "-")}-data-objects.json`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  const filteredNodes = nodes.filter(
    ({ businessFunction, capability }) =>
      (functionFilter === "all" || businessFunction.id === functionFilter) &&
      (capabilityFilter === "all" || capability.id === capabilityFilter),
  );
  const availableCapabilities =
    functionFilter === "all"
      ? sector.functions.flatMap(
          (businessFunction) => businessFunction.capabilities,
        )
      : (sector.functions.find(
          (businessFunction) => businessFunction.id === functionFilter,
        )?.capabilities ?? []);
  return (
    <section className="explorer-panel">
      <div className="process-map-filters">
        <label>
          <span className="label">Primary Function</span>
          <select
            value={functionFilter}
            onChange={(event) => {
              setFunctionFilter(event.target.value);
              setCapabilityFilter("all");
            }}
          >
            <option value="all">All functions</option>
            {sector.functions.map((businessFunction) => (
              <option value={businessFunction.id} key={businessFunction.id}>
                {businessFunction.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          <span className="label">Primary Capability</span>
          <select
            value={capabilityFilter}
            onChange={(event) => setCapabilityFilter(event.target.value)}
          >
            <option value="all">All capabilities</option>
            {availableCapabilities.map((capability) => (
              <option value={capability.id} key={capability.id}>
                {capability.name}
              </option>
            ))}
          </select>
        </label>
      </div>
      <div className="process-tree" aria-label="Process Map">
        {stages.map((stage) => {
          const stageNodes = filteredNodes.filter(
            ({ process }) => !process.stageId || process.stageId === stage.id,
          );
          const streamIds = Array.from(
            new Set(
              stageNodes.map(
                ({ process }) => process.parentProcessId ?? process.id,
              ),
            ),
          );
          const stageCollapsed = collapsedStages.has(stage.id);
          const primaryFunctions = Array.from(
            new Set(
              stageNodes.map(({ businessFunction }) => businessFunction.name),
            ),
          );
          const primaryCapabilities = Array.from(
            new Set(stageNodes.map(({ capability }) => capability.name)),
          );
          const supportingFunctions = Array.from(
            new Set(
              stageNodes.flatMap(
                ({ process }) => process.contributingFunctions ?? [],
              ),
            ),
          );
          const supportingCapabilities = Array.from(
            new Set(
              stageNodes.flatMap(
                ({ process }) => process.contributingCapabilities ?? [],
              ),
            ),
          );
          const stageDataObjects = Array.from(
            new Map(
              stageNodes
                .flatMap((node) => dataObjectsForNode(node))
                .map((entity) => [entity.id, entity]),
            ).values(),
          );
          return (
            <section className="tree-stage" key={stage.id}>
              <div
                className="tree-node tree-stage-node"
                role="button"
                tabIndex={0}
                onClick={() => toggleStage(stage.id, streamIds)}
                onKeyDown={(event) => {
                  if (event.key === "Enter" || event.key === " ")
                    toggleStage(stage.id, streamIds);
                }}
                aria-expanded={!stageCollapsed}
              >
                <span className="tree-index">
                  {String(stage.order).padStart(2, "0")}
                </span>
                <span>
                  <span className="label">L1 Value Chain</span>
                  <strong>{stage.name}</strong>
                  <span className="objective-row">
                    <b>Objective:</b>
                    <span>{stage.description}</span>
                  </span>
                  <span className="taxonomy-summary">
                    <b>Primary Functions:</b>{" "}
                    {primaryFunctions.join(", ") || "None in filter"}
                    <b>Primary Capabilities:</b>{" "}
                    {primaryCapabilities.join(", ") || "None in filter"}
                    <b>Supporting Functions:</b>{" "}
                    {supportingFunctions.join(", ") || "None identified"}
                    <b>Supporting Capabilities:</b>{" "}
                    {supportingCapabilities.join(", ") || "None identified"}
                  </span>
                  <span className="data-object-list l1-data-objects">
                    <span className="label">
                      Data Objects{" "}
                      <button
                        className="json-download"
                        type="button"
                        title="Download L1 Data Objects JSON"
                        onClick={(event) => {
                          event.stopPropagation();
                          downloadDataObjectsJson(
                            "L1",
                            stage.name,
                            stageDataObjects,
                          );
                        }}
                      >
                        <Download size={13} />
                      </button>
                    </span>
                    {stageDataObjects.length ? (
                      stageDataObjects.map((entity) => (
                        <button
                          type="button"
                          key={entity.id}
                          onClick={(event) => {
                            event.stopPropagation();
                            setSelectedDataObject(entity);
                            setSelectedDataObjectLevel("l1");
                            setSelectedDataObjectNode(stage.id);
                          }}
                        >
                          {entity.name.replace(/_/g, " ")}
                        </button>
                      ))
                    ) : (
                      <small>None mapped</small>
                    )}
                  </span>
                </span>
                <ChevronDown size={18} />
              </div>
              {selectedDataObjectLevel === "l1" &&
                selectedDataObjectNode === stage.id &&
                selectedDataObject &&
                stageDataObjects.some(
                  (entity) => entity.id === selectedDataObject.id,
                ) && (
                  <div className="data-object-detail l1-data-object-detail">
                    <div>
                      <span className="label">Selected Data Object</span>
                      <strong>
                        {selectedDataObject.name.replace(/_/g, " ")}
                      </strong>
                      <small>{selectedDataObject.description}</small>
                    </div>
                    <div>
                      <span className="label">Attributes</span>
                      <div className="detail-chips">
                        {selectedDataObject.attributes.map((attribute) => (
                          <code key={attribute}>{attribute}</code>
                        ))}
                      </div>
                    </div>
                    <div>
                      <span className="label">Connections</span>
                      <div className="detail-connections">
                        {selectedDataObject.relationships.length ? (
                          selectedDataObject.relationships.map(
                            (relationship) => (
                              <span key={relationship}>{relationship}</span>
                            ),
                          )
                        ) : (
                          <small>No relationships mapped</small>
                        )}
                      </div>
                    </div>
                  </div>
                )}
              {!stageCollapsed && (
                <div className="tree-children">
                  {streamIds.map((streamId, streamIndex) => {
                    const streamNodes = stageNodes.filter(
                      ({ process }) =>
                        (process.parentProcessId ?? process.id) === streamId,
                    );
                    const first = streamNodes[0];
                    const dataObjects = Array.from(
                      new Map(
                        streamNodes
                          .flatMap((node) => dataObjectsForNode(node))
                          .map((entity) => [entity.id, entity]),
                      ).values(),
                    );
                    const streamKey = `${stage.id}-${streamId}`;
                    const streamCollapsed = collapsedStreams.has(streamKey);
                    return (
                      <article className="tree-work-stream" key={streamKey}>
                        <div
                          className="tree-node tree-stream-node"
                          role="button"
                          tabIndex={0}
                          onClick={() =>
                            toggleCollapsed(setCollapsedStreams, streamKey)
                          }
                          onKeyDown={(event) => {
                            if (event.key === "Enter" || event.key === " ")
                              toggleCollapsed(setCollapsedStreams, streamKey);
                          }}
                          aria-expanded={!streamCollapsed}
                        >
                          <span className="tree-index">
                            {String(streamIndex + 1).padStart(2, "0")}
                          </span>
                          <span>
                            <span className="label">L2 Work Stream</span>
                            <strong>
                              {first.process.parentProcessName ??
                                first.process.name}
                            </strong>
                            <span className="objective-row compact">
                              <b>Objective:</b>
                              <span>
                                {first.process.parentProcessDescription ??
                                  "Not specified"}
                              </span>
                            </span>
                            <small>
                              {first.businessFunction.name} /{" "}
                              {first.capability.name}
                            </small>
                          </span>
                          <span className="data-object-list">
                            <span className="label">
                              Data Objects{" "}
                              <button
                                className="json-download"
                                type="button"
                                title="Download L2 Data Objects JSON"
                                onClick={(event) => {
                                  event.stopPropagation();
                                  downloadDataObjectsJson(
                                    "L2",
                                    first.process.parentProcessName ??
                                      first.process.name,
                                    dataObjects,
                                  );
                                }}
                              >
                                <Download size={13} />
                              </button>
                            </span>
                            {dataObjects.length ? (
                              dataObjects.map((entity) => (
                                <button
                                  type="button"
                                  key={entity.id}
                                  onClick={(event) => {
                                    event.stopPropagation();
                                    setSelectedDataObject(entity);
                                    setSelectedDataObjectLevel("l2");
                                    setSelectedDataObjectNode(streamKey);
                                  }}
                                >
                                  {entity.name.replace(/_/g, " ")}
                                </button>
                              ))
                            ) : (
                              <small>None mapped</small>
                            )}
                          </span>
                          <ChevronDown size={18} />
                        </div>
                        {selectedDataObjectLevel === "l2" &&
                          selectedDataObjectNode === streamKey &&
                          selectedDataObject &&
                          dataObjects.some(
                            (entity) => entity.id === selectedDataObject.id,
                          ) && (
                            <div className="data-object-detail">
                              <div>
                                <span className="label">
                                  Selected Data Object
                                </span>
                                <strong>
                                  {selectedDataObject.name.replace(/_/g, " ")}
                                </strong>
                                <small>{selectedDataObject.description}</small>
                              </div>
                              <div>
                                <span className="label">Attributes</span>
                                <div className="detail-chips">
                                  {selectedDataObject.attributes.map(
                                    (attribute) => (
                                      <code key={attribute}>{attribute}</code>
                                    ),
                                  )}
                                </div>
                              </div>
                              <div>
                                <span className="label">Connections</span>
                                <div className="detail-connections">
                                  {selectedDataObject.relationships.length ? (
                                    selectedDataObject.relationships.map(
                                      (relationship) => (
                                        <span key={relationship}>
                                          {relationship}
                                        </span>
                                      ),
                                    )
                                  ) : (
                                    <small>No relationships mapped</small>
                                  )}
                                </div>
                              </div>
                            </div>
                          )}
                        {!streamCollapsed && (
                          <div className="tree-children activity-list">
                            {streamNodes.map((node, index) => {
                              const process = node.process;
                              const processDataObjects =
                                dataObjectsForNode(node);
                              const open = expandedProcess === process.id;
                              return (
                                <div
                                  className={
                                    open ? "activity-row open" : "activity-row"
                                  }
                                  key={process.id}
                                >
                                  <button
                                    className="tree-node activity-node"
                                    type="button"
                                    onClick={() => onProcessToggle(process.id)}
                                    aria-expanded={open}
                                  >
                                    <span className="tree-index">
                                      {String(index + 1).padStart(2, "0")}
                                    </span>
                                    <span>
                                      <span className="label">L3 Activity</span>
                                      <strong>{process.name}</strong>
                                      <span className="objective-row compact">
                                        <b>Objective:</b>
                                        <span>{process.description}</span>
                                      </span>
                                    </span>
                                    <ChevronDown size={18} />
                                  </button>
                                  <div className="l3-data-objects">
                                    <span className="label">
                                      Data Objects{" "}
                                      <button
                                        className="json-download"
                                        type="button"
                                        title="Download L3 Data Objects JSON"
                                        onClick={(event) => {
                                          event.stopPropagation();
                                          downloadDataObjectsJson(
                                            "L3",
                                            process.name,
                                            processDataObjects,
                                          );
                                        }}
                                      >
                                        <Download size={13} />
                                      </button>
                                    </span>
                                    {processDataObjects.length ? (
                                      processDataObjects.map((entity) => (
                                        <button
                                          type="button"
                                          key={entity.id}
                                          onClick={(event) => {
                                            event.stopPropagation();
                                            setSelectedDataObject(entity);
                                            setSelectedDataObjectLevel("l3");
                                            setSelectedDataObjectNode(
                                              process.id,
                                            );
                                          }}
                                        >
                                          {entity.name.replace(/_/g, " ")}
                                        </button>
                                      ))
                                    ) : (
                                      <small>None mapped</small>
                                    )}
                                  </div>
                                  {selectedDataObjectLevel === "l3" &&
                                    selectedDataObjectNode === process.id &&
                                    selectedDataObject &&
                                    processDataObjects.some(
                                      (entity) =>
                                        entity.id === selectedDataObject.id,
                                    ) && (
                                      <div className="data-object-detail l3-data-object-detail">
                                        <div>
                                          <span className="label">
                                            Selected Data Object
                                          </span>
                                          <strong>
                                            {selectedDataObject.name.replace(
                                              /_/g,
                                              " ",
                                            )}
                                          </strong>
                                          <small>
                                            {selectedDataObject.description}
                                          </small>
                                        </div>
                                        <div>
                                          <span className="label">
                                            Attributes
                                          </span>
                                          <div className="detail-chips">
                                            {selectedDataObject.attributes.map(
                                              (attribute) => (
                                                <code key={attribute}>
                                                  {attribute}
                                                </code>
                                              ),
                                            )}
                                          </div>
                                        </div>
                                        <div>
                                          <span className="label">
                                            Connections
                                          </span>
                                          <div className="detail-connections">
                                            {selectedDataObject.relationships
                                              .length ? (
                                              selectedDataObject.relationships.map(
                                                (relationship) => (
                                                  <span key={relationship}>
                                                    {relationship}
                                                  </span>
                                                ),
                                              )
                                            ) : (
                                              <small>
                                                No relationships mapped
                                              </small>
                                            )}
                                          </div>
                                        </div>
                                      </div>
                                    )}
                                  {open && (
                                    <div className="process-detail">
                                      <div className="io-strip">
                                        <div>
                                          <span>Inputs</span>
                                          {process.inputs.join(" + ") ||
                                            "Not specified"}
                                        </div>
                                        <ArrowRight size={18} />
                                        <div>
                                          <span>Outputs</span>
                                          {process.outputs.join(" + ") ||
                                            "Not specified"}
                                        </div>
                                      </div>
                                      <div className="step-flow">
                                        {process.steps.map(
                                          (step, stepIndex) => (
                                            <div
                                              className="flow-pair"
                                              key={`${process.id}-${stepIndex}`}
                                            >
                                              <div className="step-node">
                                                <span>{stepIndex + 1}</span>
                                                <strong>{step}</strong>
                                              </div>
                                              {stepIndex <
                                                process.steps.length - 1 && (
                                                <ArrowRight
                                                  className="flow-arrow"
                                                  size={17}
                                                />
                                              )}
                                            </div>
                                          ),
                                        )}
                                      </div>
                                    </div>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </article>
                    );
                  })}
                </div>
              )}
            </section>
          );
        })}
      </div>
    </section>
  );
}

type LegacyProcessExplorerProps = {
  sector: Sector;
  businessFunction: Sector["functions"][number];
  capabilityId: string;
  expandedProcess: string | null;
  onCapabilityChange: (id: string) => void;
  onProcessToggle: (id: string) => void;
};
function LegacyProcessExplorer({
  sector,
  businessFunction,
  capabilityId,
  expandedProcess,
  onCapabilityChange,
  onProcessToggle,
}: LegacyProcessExplorerProps) {
  const capability =
    businessFunction.capabilities.find((item) => item.id === capabilityId) ??
    businessFunction.capabilities[0];
  const [stageId, setStageId] = useState(sector.stages?.[0]?.id ?? "");
  const stageProcesses = capability.processes.filter(
    (process) => !process.stageId || process.stageId === stageId,
  );
  const workStreams = Array.from(
    new Map(
      stageProcesses.map((process) => [
        process.parentProcessId ?? process.id,
        process.parentProcessName ?? process.name,
      ]),
    ).entries(),
  );
  /*
  return <section className="explorer-panel"><div className="stage-rail" role="tablist" aria-label="Value Chain"><span className="label">L1 Value Chain</span>{sector.stages?.map((stage) => <button className={stage.id === stageId ? 'active' : ''} type="button" key={stage.id} onClick={() => { setStageId(stage.id); setExpandedProcess(null) }}>{String(stage.order).padStart(2, '0')} {stage.name}</button>)}</div><div className="capability-rail" role="tablist" aria-label="Capabilities">{businessFunction.capabilities.map((item) => <button className={item.id === capability.id ? 'active' : ''} type="button" key={item.id} onClick={() => onCapabilityChange(item.id)}>{item.name}<span>{item.processes.length}</span></button>)}</div><div className="capability-summary"><div><span className="label">Function / Capability</span><h2>{businessFunction.name} / {capability.name}</h2><p>{capability.description}</p></div><div className="system-list"><span className="label">Systems of record</span>{capability.systems.map((system) => <span key={system}>{system}</span>)}</div></div><div className="process-list">{workStreams.map(([workStreamId, workStreamName], streamIndex) => { const activities = stageProcesses.filter((process) => (process.parentProcessId ?? process.id) === workStreamId); const dataObjects = sector.entities.filter((entity) => entity.capabilityId === capability.id); return <article className="work-stream-row" key={workStreamId}><div className="work-stream-header"><span className="process-number">{String(streamIndex + 1).padStart(2, '0')}</span><div><span className="label">L2 Work Stream</span><strong>{workStreamName}</strong><small>{businessFunction.name} / {capability.name}</small></div><div className="data-object-list"><span className="label">Data Objects</span>{dataObjects.length ? dataObjects.map((entity) => <span key={entity.id}>{entity.name.replace(/_/g, ' ')}</span>) : <small>None mapped</small>}</div></div><div className="activity-list">{activities.map((process, index) => { const open = expandedProcess === process.id; return <div className={open ? 'activity-row open' : 'activity-row'} key={process.id}><button className="process-summary" type="button" onClick={() => onProcessToggle(process.id)} aria-expanded={open}><span className="process-number">{String(index + 1).padStart(2, '0')}</span><span className="process-title"><small>L3 Activity · {process.id}</small><strong>{process.name}</strong><em>{process.description}</em></span><ChevronDown size={19} /></button>{open && <div className="process-detail"><div className="io-strip"><div><span>Inputs</span>{process.inputs.join(' + ') || 'Not specified'}</div><ArrowRight size={18} /><div><span>Outputs</span>{process.outputs.join(' + ') || 'Not specified'}</div></div><div className="step-flow">{process.steps.map((step, stepIndex) => <div className="flow-pair" key={`${process.id}-${stepIndex}`}><div className="step-node"><span>{stepIndex + 1}</span><strong>{step}</strong></div>{stepIndex < process.steps.length - 1 && <ArrowRight className="flow-arrow" size={17} />}</div>)}</div></div>}</div>)}</div></article>})}</div></section>
  */
  return (
    <section className="explorer-panel">
      <div className="stage-rail" role="tablist" aria-label="Value Chain">
        <span className="label">L1 Value Chain</span>
        {sector.stages?.map((stage) => (
          <button
            className={stage.id === stageId ? "active" : ""}
            type="button"
            key={stage.id}
            onClick={() => {
              setStageId(stage.id);
              if (expandedProcess) onProcessToggle(expandedProcess);
            }}
          >
            <span>{String(stage.order).padStart(2, "0")}</span>
            <strong>{stage.name}</strong>
            <ChevronDown size={16} />
          </button>
        ))}
      </div>
      <div className="capability-rail" role="tablist" aria-label="Capabilities">
        {businessFunction.capabilities.map((item) => (
          <button
            className={item.id === capability.id ? "active" : ""}
            type="button"
            key={item.id}
            onClick={() => onCapabilityChange(item.id)}
          >
            {item.name}
            <span>{item.processes.length}</span>
          </button>
        ))}
      </div>
      <div className="capability-summary">
        <div>
          <span className="label">Function / Capability</span>
          <h2>
            {businessFunction.name} / {capability.name}
          </h2>
          <p>{capability.description}</p>
        </div>
        <div className="system-list">
          <span className="label">Systems of record</span>
          {capability.systems.map((system) => (
            <span key={system}>{system}</span>
          ))}
        </div>
      </div>
      <div className="process-list">
        {workStreams.map(([workStreamId, workStreamName], streamIndex) => {
          const activities = stageProcesses.filter(
            (process) =>
              (process.parentProcessId ?? process.id) === workStreamId,
          );
          const dataObjects = sector.entities.filter(
            (entity) => entity.capabilityId === capability.id,
          );
          return (
            <article className="work-stream-row" key={workStreamId}>
              <div className="work-stream-header">
                <span className="process-number">
                  {String(streamIndex + 1).padStart(2, "0")}
                </span>
                <div>
                  <span className="label">L2 Work Stream</span>
                  <strong>{workStreamName}</strong>
                  <small>
                    {businessFunction.name} / {capability.name}
                  </small>
                </div>
                <div className="data-object-list">
                  <span className="label">Data Objects</span>
                  {dataObjects.length ? (
                    dataObjects.map((entity) => (
                      <span key={entity.id}>
                        {entity.name.replace(/_/g, " ")}
                      </span>
                    ))
                  ) : (
                    <small>None mapped</small>
                  )}
                </div>
              </div>
              <div className="activity-list">
                {activities.map((process, index) => {
                  const open = expandedProcess === process.id;
                  return (
                    <div
                      className={open ? "activity-row open" : "activity-row"}
                      key={process.id}
                    >
                      <button
                        className="process-summary"
                        type="button"
                        onClick={() => onProcessToggle(process.id)}
                        aria-expanded={open}
                      >
                        <span className="process-number">
                          {String(index + 1).padStart(2, "0")}
                        </span>
                        <span className="process-title">
                          <small>L3 Activity</small>
                          <strong>{process.name}</strong>
                          <em>{process.description}</em>
                        </span>
                        <ChevronDown size={19} />
                      </button>
                      {open && (
                        <div className="process-detail">
                          <div className="io-strip">
                            <div>
                              <span>Inputs</span>
                              {process.inputs.join(" + ") || "Not specified"}
                            </div>
                            <ArrowRight size={18} />
                            <div>
                              <span>Outputs</span>
                              {process.outputs.join(" + ") || "Not specified"}
                            </div>
                          </div>
                          <div className="step-flow">
                            {process.steps.map((step, stepIndex) => (
                              <div
                                className="flow-pair"
                                key={`${process.id}-${stepIndex}`}
                              >
                                <div className="step-node">
                                  <span>{stepIndex + 1}</span>
                                  <strong>{step}</strong>
                                </div>
                                {stepIndex < process.steps.length - 1 && (
                                  <ArrowRight
                                    className="flow-arrow"
                                    size={17}
                                  />
                                )}
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </article>
          );
        })}
      </div>
    </section>
  );
}

void LegacyProcessExplorer;

type OntologyExplorerProps = {
  sector: Sector;
  search: string;
  selectedEntity: DataEntity | null;
  onSearch: (value: string) => void;
  onSelect: (entity: DataEntity | null) => void;
};
const SYSTEM_COLORS: Record<string, string> = {
  pd: "#6B5FBF",
  qms: "#1E8C89",
  erp: "#2E7DBF",
  mes: "#C24747",
  lims: "#B98A2E",
  rims: "#7A5FBF",
  clin: "#3D9970",
  pv: "#B33D6B",
};
const DEFAULT_NODE_COLOR = "#8891A0";

function OntologyExplorer({
  sector,
  search,
  selectedEntity,
  onSearch,
  onSelect,
}: OntologyExplorerProps) {
  const [functionFilter, setFunctionFilter] = useState("all");
  const [capabilityFilter, setCapabilityFilter] = useState("all");
  const [l1Filter, setL1Filter] = useState("all");
  const [l2Filter, setL2Filter] = useState("all");
  const [l4Filter, setL4Filter] = useState("all");
  const [ontologyMode, setOntologyMode] = useState<"network" | "list" | "graph">("network");
  const [graphScale, setGraphScale] = useState(1);
  const [graphPan, setGraphPan] = useState({ x: 0, y: 0 });
  const [dragging, setDragging] = useState(false);
  const graphDrag = useRef({ x: 0, y: 0, panX: 0, panY: 0 });
  const query = search.trim().toLowerCase();
  const unique = (values: Array<string | undefined>) =>
    Array.from(new Set(values.filter(Boolean) as string[])).sort();
  const functions = unique(
    sector.entities.map((entity) => entity.primaryFunction),
  );
  const capabilities = unique(
    sector.entities
      .filter(
        (entity) =>
          functionFilter === "all" || entity.primaryFunction === functionFilter,
      )
      .map((entity) => entity.primaryCapability),
  );
  const l1Values = unique(sector.entities.map((entity) => entity.l1Name));
  const l2Values = unique(
    sector.entities
      .filter((entity) => l1Filter === "all" || entity.l1Name === l1Filter)
      .map((entity) => entity.l2Name),
  );
  const l4Values = unique(
    sector.entities
      .filter((entity) => l2Filter === "all" || entity.l2Name === l2Filter)
      .map((entity) => entity.l3Name),
  );
  const entities = sector.entities.filter(
    (entity) =>
      (functionFilter === "all" || entity.primaryFunction === functionFilter) &&
      (capabilityFilter === "all" ||
        entity.primaryCapability === capabilityFilter) &&
      (l1Filter === "all" || entity.l1Name === l1Filter) &&
      (l2Filter === "all" || entity.l2Name === l2Filter) &&
      (l4Filter === "all" || entity.l3Name === l4Filter) &&
      (!query ||
        [entity.name, entity.system, entity.description, ...entity.attributes]
          .join(" ")
          .toLowerCase()
          .includes(query)),
  );
  const entityIds = new Set(entities.map((entity) => entity.name));
  const relationships = (sector.entityRelationships ?? []).filter(
    (relationship) =>
      entityIds.has(relationship.from) && entityIds.has(relationship.to),
  );
  const nodeWidth = 240;
  const nodeHeight = 146;
  const gridColumns = 3;
  const gridGapX = 72;
  const gridGapY = 64;
  const graphWidth = gridColumns * (nodeWidth + gridGapX) + 80;
  const graphHeight = Math.max(700, Math.ceil(entities.length / gridColumns) * (nodeHeight + gridGapY) + 100);
  const positions = new Map(entities.map((entity, index) => [entity.name, { x: (index % gridColumns) * (nodeWidth + gridGapX) + 40, y: Math.floor(index / gridColumns) * (nodeHeight + gridGapY) + 70 }]));

  // Graph view: small colored dots clustered by system (rather than a
  // dense grid of full attribute cards) -- grouping same-type nodes
  // together is what keeps this readable instead of a tangled mess.
  const dotRadius = 14;
  const systemGroups = new Map<string, typeof entities>();
  entities.forEach((entity) => {
    const key = entity.system || "other";
    const list = systemGroups.get(key) ?? [];
    list.push(entity);
    systemGroups.set(key, list);
  });
  const systemKeys = Array.from(systemGroups.keys()).sort();
  const clusterCount = systemKeys.length;
  const maxClusterSize = Math.max(1, ...systemKeys.map((k) => systemGroups.get(k)!.length));
  const clusterRadius = Math.max(46, Math.min(120, 20 * Math.sqrt(maxClusterSize)));
  const bigRadius = clusterCount <= 1 ? 0 : Math.max(150, clusterCount * 42);
  const dotCanvasSize = {
    width: 2 * (bigRadius + clusterRadius) + 200,
    height: 2 * (bigRadius + clusterRadius) + 200,
  };
  const dotCenter = { x: dotCanvasSize.width / 2, y: dotCanvasSize.height / 2 };
  const dotPositions = new Map<string, { x: number; y: number }>();
  systemKeys.forEach((key, clusterIndex) => {
    const clusterAngle = clusterCount <= 1 ? 0 : (clusterIndex / clusterCount) * 2 * Math.PI - Math.PI / 2;
    const clusterCenter = clusterCount <= 1
      ? dotCenter
      : { x: dotCenter.x + bigRadius * Math.cos(clusterAngle), y: dotCenter.y + bigRadius * Math.sin(clusterAngle) };
    const members = systemGroups.get(key)!;
    members.forEach((entity, memberIndex) => {
      const innerRadius = members.length <= 1 ? 0 : clusterRadius;
      const angle = members.length <= 1 ? 0 : (memberIndex / members.length) * 2 * Math.PI;
      dotPositions.set(entity.name, {
        x: clusterCenter.x + innerRadius * Math.cos(angle),
        y: clusterCenter.y + innerRadius * Math.sin(angle),
      });
    });
  });

  function zoomGraph(delta: number) {
    setGraphScale((current) =>
      Math.min(2, Math.max(0.55, Number((current + delta).toFixed(2)))),
    );
  }
  function handleGraphWheel(event: React.WheelEvent<HTMLDivElement>) {
    event.preventDefault();
    zoomGraph(event.deltaY < 0 ? 0.1 : -0.1);
  }
  function startGraphDrag(event: React.PointerEvent<HTMLDivElement>) {
    if ((event.target as HTMLElement).closest("button")) return;
    graphDrag.current = {
      x: event.clientX,
      y: event.clientY,
      panX: graphPan.x,
      panY: graphPan.y,
    };
    setDragging(true);
    event.currentTarget.setPointerCapture(event.pointerId);
  }
  function moveGraph(event: React.PointerEvent<HTMLDivElement>) {
    if (!dragging) return;
    setGraphPan({
      x: graphDrag.current.panX + event.clientX - graphDrag.current.x,
      y: graphDrag.current.panY + event.clientY - graphDrag.current.y,
    });
  }
  function downloadFilteredOntologyJson() {
    const names = new Set(entities.map((entity) => entity.name));
    const connections = (sector.entityRelationships ?? []).filter(
      (relationship) => names.has(relationship.from) && names.has(relationship.to),
    );
    const payload = {
      level: "Filtered",
      filters: { functionFilter, capabilityFilter, l1Filter, l2Filter, l3Filter: l4Filter, search },
      dataObjects: entities.map((entity) => ({
        name: entity.name,
        attributes: Object.fromEntries(entity.attributes.map((attribute) => [attribute, ""])),
        connections: connections.filter((connection) => connection.from === entity.name || connection.to === entity.name),
      })),
    };
    const url = URL.createObjectURL(new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${sector.id}-filtered-data-objects.json`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  return (
    <section className="ontology-layout">
      <div className="ontology-main">
        <div className="ontology-toolbar">
          <div>
            <span className="label">Governed data model</span>
            <h2>Connected Data Objects</h2>
          </div>
          <button className="ontology-json-download" type="button" onClick={downloadFilteredOntologyJson}><Download size={15} /> Download JSON</button>
          <label className="search-box">
            <Search size={17} />
            <input
              value={search}
              onChange={(event) => onSearch(event.target.value)}
              placeholder="Search entities or attributes"
            />
          </label>
        </div>
        <div className="ontology-filters">
          {[
            ["Business Function", functionFilter, setFunctionFilter, functions],
            ["Capability", capabilityFilter, setCapabilityFilter, capabilities],
            ["L1 Value Chain", l1Filter, setL1Filter, l1Values],
            ["L2 Work Stream", l2Filter, setL2Filter, l2Values],
            ["L3 Activity", l4Filter, setL4Filter, l4Values],
          ].map(([label, value, setter, values]) => (
            <label key={label as string}>
              <span className="label">{label as string}</span>
              <select
                value={value as string}
                onChange={(event) => {
                  (setter as (value: string) => void)(event.target.value);
                  if (label === "Business Function") setCapabilityFilter("all");
                }}
              >
                <option value="all">All</option>
                {(values as string[]).map((option) => (
                  <option value={option} key={option}>
                    {option}
                  </option>
                ))}
              </select>
            </label>
          ))}
        </div>
        <div className="ontology-mode-tabs" role="tablist" aria-label="Data Object view">
          <button className={ontologyMode === "network" ? "active" : ""} type="button" onClick={() => setOntologyMode("network")}>Network view</button>
          <button className={ontologyMode === "graph" ? "active" : ""} type="button" onClick={() => setOntologyMode("graph")}>Graph view</button>
          <button className={ontologyMode === "list" ? "active" : ""} type="button" onClick={() => setOntologyMode("list")}>List view</button>
        </div>
        {ontologyMode === "list" && <div className="ontology-list-view">{entities.map((entity) => <button className={selectedEntity?.id === entity.id ? "ontology-list-item active" : "ontology-list-item"} type="button" key={entity.id} onClick={() => onSelect(entity)}><span className="entity-icon"><Boxes size={17} /></span><span><strong>{entity.name.replace(/_/g, " ")}</strong><small>{entity.description}</small><em>{entity.primaryFunction ?? ""} / {entity.primaryCapability ?? ""}</em></span><span className="ontology-list-count">{entity.attributes.length} attributes<br />{entity.relationships.length} connections</span></button>)}</div>}
        {(ontologyMode === "network" || ontologyMode === "graph") && entities.length ? (
          <div className="ontology-graph-shell">
            <div className="graph-controls">
              <button
                type="button"
                title="Zoom out"
                onClick={() => zoomGraph(-0.1)}
              >
                −
              </button>
              <span>{Math.round(graphScale * 100)}%</span>
              <button
                type="button"
                title="Zoom in"
                onClick={() => zoomGraph(0.1)}
              >
                +
              </button>
              <button
                type="button"
                title="Reset view"
                onClick={() => {
                  setGraphScale(1);
                  setGraphPan({ x: 0, y: 0 });
                }}
              >
                Reset
              </button>
            </div>
            <div
              className={
                dragging ? "ontology-graph dragging" : "ontology-graph"
              }
              style={{ height: 520 }}
              onWheel={handleGraphWheel}
              onPointerDown={startGraphDrag}
              onPointerMove={moveGraph}
              onPointerUp={() => setDragging(false)}
              onPointerCancel={() => setDragging(false)}
            >
              <div
                className="ontology-graph-content"
                style={{
                  width: ontologyMode === "graph" ? dotCanvasSize.width : graphWidth,
                  height: ontologyMode === "graph" ? dotCanvasSize.height : graphHeight,
                  transform: `translate(${graphPan.x}px, ${graphPan.y}px) scale(${graphScale})`,
                }}
              >
                {ontologyMode === "graph" ? (
                  <>
                    <svg className="ontology-edges" width={dotCanvasSize.width} height={dotCanvasSize.height}>
                      <defs>
                        <marker id="graph-dot-arrow" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto" markerUnits="strokeWidth">
                          <path d="M0,0 L6,3 L0,6 z" fill="#aab0ba" />
                        </marker>
                      </defs>
                      {relationships.map((relationship, relationshipIndex) => {
                        const from = dotPositions.get(relationship.from);
                        const to = dotPositions.get(relationship.to);
                        if (!from || !to) return null;
                        const midX = (from.x + to.x) / 2;
                        const midY = (from.y + to.y) / 2;
                        const dx = to.x - from.x;
                        const dy = to.y - from.y;
                        const length = Math.max(1, Math.sqrt(dx * dx + dy * dy));
                        const curveOffset = ((relationshipIndex % 5) - 2) * 12;
                        const controlX = midX - (dy / length) * curveOffset;
                        const controlY = midY + (dx / length) * curveOffset;
                        const path = `M ${from.x} ${from.y} Q ${controlX} ${controlY} ${to.x} ${to.y}`;
                        return (
                          <path
                            key={`${relationship.from}-${relationship.to}-${relationship.relationship}`}
                            d={path}
                            fill="none"
                            stroke="#c7ccd3"
                            strokeWidth={1.3}
                            opacity={0.55}
                            markerEnd="url(#graph-dot-arrow)"
                          />
                        );
                      })}
                    </svg>
                    {entities.map((entity) => {
                      const position = dotPositions.get(entity.name)!;
                      const color = SYSTEM_COLORS[entity.system?.toLowerCase()] ?? DEFAULT_NODE_COLOR;
                      const isActive = selectedEntity?.id === entity.id;
                      return (
                        <button
                          key={entity.id}
                          type="button"
                          className={isActive ? "ontology-dot-node active" : "ontology-dot-node"}
                          title={entity.name.replace(/_/g, " ")}
                          data-label={entity.name.replace(/_/g, " ")}
                          style={{
                            left: position.x - dotRadius,
                            top: position.y - dotRadius,
                            width: dotRadius * 2,
                            height: dotRadius * 2,
                            background: color,
                            borderColor: isActive ? "#1F2430" : "#ffffff",
                          }}
                          onClick={() => onSelect(entity)}
                        />
                      );
                    })}
                    <div className="ontology-graph-legend">
                      {systemKeys.map((key) => (
                        <div key={key} className="ontology-graph-legend-item">
                          <span
                            className="ontology-graph-legend-dot"
                            style={{ background: SYSTEM_COLORS[key.toLowerCase()] ?? DEFAULT_NODE_COLOR }}
                          />
                          {key.toUpperCase()}
                        </div>
                      ))}
                    </div>
                  </>
                ) : (
                  <>
                <svg
                  className="ontology-edges"
                  width={graphWidth}
                  height={graphHeight}
                >
                  <defs>
                    <marker id="relationship-arrow" markerWidth="7" markerHeight="7" refX="6" refY="3.5" orient="auto" markerUnits="strokeWidth">
                      <path d="M0,0 L7,3.5 L0,7 z" />
                    </marker>
                  </defs>
                  {relationships.map((relationship, relationshipIndex) => {
                    const from = positions.get(relationship.from);
                    const to = positions.get(relationship.to);
                    if (!from || !to) return null;
                    const forward = to.x >= from.x;
                    const x1 = from.x + (forward ? nodeWidth : 0);
                    const y1 = from.y + nodeHeight / 2;
                    const x2 = to.x + (forward ? 0 : nodeWidth);
                    const y2 = to.y + nodeHeight / 2;
                    const label = relationship.relationship
                      .replace(/_/g, " ")
                      .toUpperCase();
                    const midX = (x1 + x2) / 2;
                    const midY = (y1 + y2) / 2;
                    const dx = x2 - x1;
                    const dy = y2 - y1;
                    const length = Math.max(1, Math.sqrt(dx * dx + dy * dy));
                    const curveOffset = ((relationshipIndex % 7) - 3) * 42;
                    const controlX = midX - (dy / length) * curveOffset;
                    const controlY = midY + (dx / length) * curveOffset;
                    const labelX = controlX;
                    const labelY = controlY - 6;
                    const path = `M ${x1} ${y1} Q ${controlX} ${controlY} ${x2} ${y2}`;
                    return (
                      <g
                        key={`${relationship.from}-${relationship.to}-${relationship.relationship}`}
                      >
                        <path className="relationship-edge" markerEnd="url(#relationship-arrow)" d={path} />
                        <rect
                          className="relationship-label-bg"
                          x={labelX - Math.max(28, label.length * 3.1)}
                          y={labelY - 10}
                          width={Math.max(56, label.length * 6.2)}
                          height={16}
                          rx={3}
                        />
                        <text x={labelX} y={labelY + 1}>
                          {label}
                        </text>
                      </g>
                    );
                  })}
                </svg>
                {entities.map((entity) => {
                  const position = positions.get(entity.name)!;
                  return (
                    <button
                      className={
                        selectedEntity?.id === entity.id
                          ? "ontology-node active"
                          : "ontology-node"
                      }
                      style={{
                        left: position.x,
                        top: position.y,
                        width: nodeWidth,
                        height: nodeHeight,
                      }}
                      type="button"
                      key={entity.id}
                      onClick={() => onSelect(entity)}
                    >
                      <span className="entity-card-icon">
                        <Boxes size={16} />
                      </span>
                      <span className="ontology-node-title">
                        {entity.name.replace(/_/g, " ")}
                      </span>
                      <span className="entity-fields">
                        {entity.attributes.slice(0, 4).map((attribute) => (
                          <span key={attribute}>
                            <code>string</code>
                            <strong>{attribute}</strong>
                          </span>
                        ))}
                      </span>
                    </button>
                  );
                })}
                  </>
                )}
              </div>
            </div>
          </div>
        ) : (
          <div className="empty-state">
            <Search size={22} />
            <strong>No matching data objects</strong>
            <span>Try broader filters or search terms.</span>
          </div>
        )}
      </div>
      <aside className="entity-detail">
        {selectedEntity &&
        entities.some((entity) => entity.id === selectedEntity.id) ? (
          <>
            <button
              className="detail-close"
              type="button"
              title="Close details"
              onClick={() => onSelect(null)}
            >
              <X size={17} />
            </button>
            <span className="overline">{selectedEntity.system} entity</span>
            <h3>{selectedEntity.name.replace(/_/g, " ")}</h3>
            <p>{selectedEntity.description}</p>
            <div className="detail-block">
              <span className="label">Attributes</span>
              {selectedEntity.attributes.map((attribute) => (
                <code key={attribute}>{attribute}</code>
              ))}
            </div>
            <div className="detail-block">
              <span className="label">Process context</span>
              <div className="relationship"><GitBranch size={14} />L1: {selectedEntity.l1Name ?? "—"}</div>
              <div className="relationship"><GitBranch size={14} />L2: {selectedEntity.l2Name ?? "—"}</div>
              <div className="relationship"><GitBranch size={14} />L3: {selectedEntity.l3Name ?? "—"}</div>
              <div className="relationship"><GitBranch size={14} />Function: {selectedEntity.primaryFunction ?? "—"}</div>
              <div className="relationship"><GitBranch size={14} />Capability: {selectedEntity.primaryCapability ?? "—"}</div>
            </div>
            <div className="detail-block">
              <span className="label">Relationships</span>
              {selectedEntity.relationships.map((relationship, index) => (
                <div className="relationship" key={`${relationship}-${index}`}>
                  <GitBranch size={14} />
                  {relationship.replace(/_/g, " ")}
                </div>
              ))}
            </div>
          </>
        ) : (
          <div className="detail-placeholder">
            <Database size={25} />
            <strong>Select an entity</strong>
            <span>Inspect its attributes and relationships here.</span>
          </div>
        )}
      </aside>
    </section>
  );
}

export default App;
