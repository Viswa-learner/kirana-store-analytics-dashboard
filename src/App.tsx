import { ChangeEvent, ReactNode, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  basketPairs, categorySummary, comparison, detectMapping, FIELD_LABELS, filterSales, forecast,
  formatISO, getRange, inventorySummary, makeSampleRows, Mapping, metrics, normalizeRows,
  InventoryOverride, Period, previousRange, productSummary, quality, RawRow, Sale, trend, TrendPoint,
} from "./lib/analytics";

type Page = "dashboard" | "upload" | "reports" | "products" | "customers" | "future";
const money = (value: number) => new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(value);
const count = (value: number) => new Intl.NumberFormat("en-IN", { maximumFractionDigits: 1 }).format(value);

const navigation: { page: Page; label: string; icon: string }[] = [
  { page: "dashboard", label: "Dashboard", icon: "home" },
  { page: "upload", label: "Upload Sales Data", icon: "upload" },
  { page: "reports", label: "Sales Reports", icon: "report" },
  { page: "products", label: "Products & Stock", icon: "box" },
  { page: "customers", label: "Customer Insights", icon: "people" },
  { page: "future", label: "Future Estimates", icon: "future" },
];

function Icon({ name, size = 18 }: { name: string; size?: number }) {
  const icons: Record<string, ReactNode> = {
    home: <><rect x="3" y="3" width="7" height="7" rx="2" /><rect x="14" y="3" width="7" height="7" rx="2" /><rect x="3" y="14" width="7" height="7" rx="2" /><rect x="14" y="14" width="7" height="7" rx="2" /></>,
    upload: <><path d="M12 16V4m0 0L7 9m5-5 5 5" /><path d="M5 14v5h14v-5" /></>,
    report: <><path d="M5 3h9l5 5v13H5zM14 3v5h5M8 13h8m-8 4h8" /></>,
    box: <><path d="m4 7 8-4 8 4-8 4-8-4Z" /><path d="m4 7v10l8 4 8-4V7M12 11v10" /></>,
    people: <><circle cx="9" cy="8" r="3" /><path d="M3 20c0-4 2-6 6-6s6 2 6 6M16 5c3 0 4 4 1.5 5.5M17 14c2.7.5 4 2.4 4 5" /></>,
    future: <><path d="M3 18c3-7 5-2 8-7s5 3 10-6" /><path d="M17 5h4v4" /></>,
    calendar: <><rect x="3" y="5" width="18" height="16" rx="2" /><path d="M16 3v4M8 3v4M3 10h18" /></>,
    check: <path d="m5 12 4 4L19 6" />,
    alert: <><path d="M12 4 3 20h18L12 4Z" /><path d="M12 9v5m0 3v.1" /></>,
    search: <><circle cx="10" cy="10" r="6" /><path d="m15 15 6 6" /></>,
    arrow: <path d="M5 12h14m-5-5 5 5-5 5" />,
    rupee: <path d="M7 5h10M7 9h10M7 5c6 0 6 8 0 8l9 7" />,
    cart: <><path d="M3 4h2l2 11h10l3-8H6" /><circle cx="9" cy="20" r="1" /><circle cx="17" cy="20" r="1" /></>,
    menu: <path d="M4 7h16M4 12h16M4 17h16" />,
    close: <path d="m6 6 12 12M18 6 6 18" />,
    download: <><path d="M12 4v12m0 0 5-5m-5 5-5-5" /><path d="M5 20h14" /></>,
    plus: <path d="M12 5v14M5 12h14" />,
  };
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{icons[name] || icons.home}</svg>;
}

function Button({ children, onClick, kind = "primary", icon, type = "button" }: { children: ReactNode; onClick?: () => void; kind?: "primary" | "secondary" | "plain"; icon?: string; type?: "button" | "submit" }) {
  return <button type={type} className={`button ${kind}`} onClick={onClick}>{icon && <Icon name={icon} size={16} />}{children}</button>;
}

function Card({ children, className = "", onClick }: { children: ReactNode; className?: string; onClick?: () => void }) {
  return <section className={`card ${className} ${onClick ? "clickable" : ""}`} onClick={onClick}>{children}</section>;
}

function PageHeading({ title, text, action }: { title: string; text: string; action?: ReactNode }) {
  return <div className="page-heading"><div><div className="page-title">{title}</div><p>{text}</p></div>{action}</div>;
}

function PeriodSelector({ period, setPeriod, customStart, customEnd, setCustomStart, setCustomEnd }: {
  period: Period; setPeriod: (period: Period) => void; customStart: string; customEnd: string;
  setCustomStart: (value: string) => void; setCustomEnd: (value: string) => void;
}) {
  const options: { key: Period; label: string }[] = [{ key: "today", label: "Today" }, { key: "week", label: "This Week" }, { key: "month", label: "This Month" }, { key: "year", label: "This Year" }, { key: "custom", label: "Custom Range" }];
  return <div className="period-wrap"><div className="period-selector">{options.map((option) => <button key={option.key} className={period === option.key ? "active" : ""} onClick={() => setPeriod(option.key)}>{option.label}</button>)}</div>{period === "custom" && <div className="custom-dates"><label>From<input type="date" value={customStart} onChange={(event) => setCustomStart(event.target.value)} /></label><span>to</span><label>Until<input type="date" value={customEnd} onChange={(event) => setCustomEnd(event.target.value)} /></label></div>}</div>;
}

function ChangeText({ value }: { value: number | null }) {
  if (value === null) return <span className="no-comparison">Not enough data for comparison</span>;
  return <span className={value >= 0 ? "change up" : "change down"}>{value >= 0 ? "↑" : "↓"} {Math.abs(value).toFixed(1)}% vs previous period</span>;
}

function Kpi({ label, value, icon, change, onClick }: { label: string; value: string; icon: string; change?: number | null; onClick?: () => void }) {
  return <Card className="kpi" onClick={onClick}><div className="kpi-top"><span>{label}</span><div><Icon name={icon} /></div></div><strong>{value}</strong>{change !== undefined && <ChangeText value={change} />}</Card>;
}

function SalesChart({ points, metric, setMetric, hasProfit, selected, setSelected, futurePoints = [] }: {
  points: TrendPoint[]; metric: "revenue" | "profit"; setMetric: (value: "revenue" | "profit") => void;
  hasProfit: boolean; selected?: TrendPoint; setSelected?: (point: TrendPoint) => void; futurePoints?: { label: string; value: number }[];
}) {
  const values = points.map((point) => metric === "profit" ? point.profit || 0 : point.revenue);
  const allValues = [...values, ...futurePoints.map((point) => point.value)];
  const max = Math.max(...allValues, 1);
  const width = 760, height = 220, pad = 20;
  const totalPoints = Math.max(points.length + futurePoints.length - 1, 1);
  const coords = values.map((value, index) => ({ x: pad + index * ((width - pad * 2) / totalPoints), y: height - pad - value / max * (height - pad * 2) }));
  const futureCoords = futurePoints.map((point, index) => ({ x: pad + (points.length + index) * ((width - pad * 2) / totalPoints), y: height - pad - point.value / max * (height - pad * 2) }));
  const actualPath = coords.map((point) => `${point.x},${point.y}`).join(" ");
  const forecastPath = [...(coords.length ? [coords[coords.length - 1]] : []), ...futureCoords].map((point) => `${point.x},${point.y}`).join(" ");
  return <div>
    <div className="chart-toolbar"><div><b>Sales trend</b><span>Click any point to see its details</span></div>{hasProfit && <div className="metric-toggle"><button className={metric === "revenue" ? "active" : ""} onClick={() => setMetric("revenue")}>Sales</button><button className={metric === "profit" ? "active" : ""} onClick={() => setMetric("profit")}>Profit</button></div>}</div>
    {points.length ? <div className="chart-box"><div className="chart-y"><span>{money(max)}</span><span>{money(max / 2)}</span><span>₹0</span></div><svg className="sales-chart" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
      {[20, 110, 200].map((y) => <line key={y} x1="20" x2="740" y1={y} y2={y} className="grid-line" />)}
      {coords.length > 1 && <polyline points={actualPath} className="actual-line" />}
      {forecastCoords(forecastPath)}
      {coords.map((point, index) => <circle key={points[index].key} cx={point.x} cy={point.y} r={selected?.key === points[index].key ? 6 : 4} className="chart-point" onClick={() => setSelected?.(points[index])}><title>{points[index].label}: {money(values[index])}</title></circle>)}
    </svg><div className="chart-labels">{[...points, ...futurePoints].filter((_, index, all) => index % Math.max(1, Math.ceil(all.length / 6)) === 0).map((point) => <span key={point.label}>{point.label}</span>)}</div></div> : <Empty title="No sales in this period" text="Choose another period or upload sales data covering these dates." />}
  </div>;
}

function forecastCoords(points: string) {
  return points ? <polyline points={points} className="forecast-line" /> : null;
}

function Empty({ title, text }: { title: string; text: string }) {
  return <div className="empty"><div><Icon name="alert" /></div><b>{title}</b><span>{text}</span></div>;
}

function DetailPanel({ point, previous }: { point: TrendPoint; previous: Sale[] }) {
  const data = metrics(point.sales);
  const best = productSummary(point.sales)[0];
  const prior = metrics(previous);
  return <div className="point-detail"><div><span>Selected period</span><b>{point.label}</b></div><div><span>Sales</span><b>{money(data.revenue)}</b></div><div><span>Transactions</span><b>{data.transactions || "Not available"}</b></div><div><span>Items sold</span><b>{count(data.units)}</b></div><div><span>Best-selling item</span><b>{best?.product || "Not available"}</b></div><div><span>Comparison</span><ChangeText value={comparison(data.revenue, prior.revenue)} /></div></div>;
}

function HorizontalBarChart({ items, selected, onSelect }: {
  items: { label: string; value: number; note: string }[];
  selected?: string;
  onSelect: (label: string) => void;
}) {
  const max = Math.max(...items.map((item) => item.value), 1);
  return <div className="horizontal-bars">{items.map((item, index) => <button key={item.label} className={selected === item.label ? "selected" : ""} onClick={() => onSelect(item.label)} title={`${item.label}: ${item.note}`}><span className="bar-rank">{index + 1}</span><div><span className="bar-name">{item.label}</span><span className="bar-track"><i style={{ width: `${item.value / max * 100}%` }} /></span></div><strong>{item.note}</strong></button>)}</div>;
}

function CategoryDonut({ items, selected, onSelect }: {
  items: { category: string; revenue: number }[];
  selected?: string;
  onSelect: (category: string) => void;
}) {
  const palette = ["#176b52", "#3d8e76", "#73b39e", "#a8d3c6", "#e0b45e", "#7792ba"];
  const total = items.reduce((sum, item) => sum + item.revenue, 0);
  let current = 0;
  const gradient = items.slice(0, 6).map((item, index) => {
    const start = current;
    current += item.revenue / (total || 1) * 100;
    return `${palette[index]} ${start}% ${current}%`;
  }).join(", ");
  const active = items.find((item) => item.category === selected) || items[0];
  return <div className="donut-chart"><button className="donut-ring" style={{ background: `conic-gradient(${gradient})` }} aria-label="Category contribution chart"><span><b>{active ? money(active.revenue) : "₹0"}</b><small>{active?.category || "No data"}</small></span></button><div className="donut-legend">{items.slice(0, 6).map((item, index) => <button key={item.category} className={selected === item.category ? "selected" : ""} onClick={() => onSelect(item.category)}><i style={{ background: palette[index] }} /><span>{item.category}</span><b>{total ? (item.revenue / total * 100).toFixed(1) : 0}%</b></button>)}</div></div>;
}

function WeekdayChart({ sales, selected, onSelect }: { sales: Sale[]; selected?: string; onSelect: (day: string) => void }) {
  const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
  const values = days.map((day) => ({ day, value: sales.filter((sale) => sale.date.toLocaleDateString("en-IN", { weekday: "long" }) === day).reduce((sum, sale) => sum + sale.quantity * sale.price, 0) }));
  const max = Math.max(...values.map((item) => item.value), 1);
  return <div className="weekday-chart">{values.map((item) => <button key={item.day} className={selected === item.day ? "selected" : ""} onClick={() => onSelect(item.day)} title={`${item.day}: ${money(item.value)}`}><strong>{money(item.value)}</strong><div><i style={{ height: `${Math.max(item.value / max * 100, 3)}%` }} /></div><span>{item.day.slice(0, 3)}</span></button>)}</div>;
}

function App() {
  const sampleRows = useMemo(() => makeSampleRows(), []);
  const [rawRows, setRawRows] = useState<RawRow[]>(sampleRows);
  const [mapping, setMapping] = useState<Mapping>(() => detectMapping(Object.keys(sampleRows[0])));
  const [dataset, setDataset] = useState("Demo Kirana Sales");
  const [isSample, setIsSample] = useState(true);
  const [page, setPage] = useState<Page>("dashboard");
  const [period, setPeriod] = useState<Period>("month");
  const [customStart, setCustomStart] = useState(formatISO(new Date(Date.now() - 29 * 86400000)));
  const [customEnd, setCustomEnd] = useState(formatISO(new Date()));
  const [inventoryOverrides, setInventoryOverrides] = useState<Record<string, InventoryOverride>>({});
  const [mobileNav, setMobileNav] = useState(false);

  const allSales = useMemo(() => normalizeRows(rawRows, mapping), [rawRows, mapping]);
  const range = useMemo(() => getRange(period, customStart, customEnd), [period, customStart, customEnd]);
  const filtered = useMemo(() => filterSales(allSales, range), [allSales, range]);
  const previous = useMemo(() => filterSales(allSales, previousRange(range)), [allSales, range]);
  const inventory = useMemo(() => inventorySummary(allSales, inventoryOverrides), [allSales, inventoryOverrides]);

  const shared = { period, setPeriod, customStart, customEnd, setCustomStart, setCustomEnd, range, filtered, previous, allSales, inventory };
  const pages: Record<Page, ReactNode> = {
    dashboard: <Dashboard {...shared} isSample={isSample} setPage={setPage} />,
    upload: <Upload rawRows={rawRows} setRawRows={setRawRows} mapping={mapping} setMapping={setMapping} dataset={dataset} setDataset={setDataset} setIsSample={setIsSample} onAnalysed={() => setPage("dashboard")} />,
    reports: <Reports {...shared} dataset={dataset} />,
    products: <Products allSales={allSales} filtered={filtered} inventory={inventory} overrides={inventoryOverrides} setOverrides={setInventoryOverrides} />,
    customers: <Customers filtered={filtered} previous={previous} rangeLabel={range.label} />,
    future: <Future allSales={allSales} inventory={inventory} />,
  };

  return <div className="app">
    <aside className={mobileNav ? "open" : ""}>
      <div className="brand"><div><Icon name="cart" /></div><span><b>KiranaBiz</b><small>Shop Assistant</small></span><button className="close-nav" onClick={() => setMobileNav(false)}><Icon name="close" /></button></div>
      <div className="nav-title">MAIN MENU</div>
      <nav>{navigation.map((item) => <button key={item.page} className={page === item.page ? "active" : ""} onClick={() => { setPage(item.page); setMobileNav(false); }}><Icon name={item.icon} /><span>{item.label}</span></button>)}</nav>
      <div className="dataset-card"><span>Using sales data</span><b>{dataset}</b><small>{allSales.length.toLocaleString("en-IN")} valid rows</small></div>
      <div className="owner"><div>RS</div><span><b>Ramesh Stores</b><small>Shop Owner</small></span></div>
    </aside>
    {mobileNav && <div className="scrim" onClick={() => setMobileNav(false)} />}
    <div className="content"><header><button className="menu" onClick={() => setMobileNav(true)}><Icon name="menu" /></button><div><span>Current data</span><b>{dataset}</b></div><div className="ready"><i /> Ready to use</div><Button kind="secondary" icon="upload" onClick={() => setPage("upload")}>Upload new data</Button></header><main>{pages[page]}</main></div>
  </div>;
}

type SharedProps = {
  period: Period; setPeriod: (period: Period) => void; customStart: string; customEnd: string;
  setCustomStart: (value: string) => void; setCustomEnd: (value: string) => void;
  range: ReturnType<typeof getRange>; filtered: Sale[]; previous: Sale[]; allSales: Sale[];
  inventory: ReturnType<typeof inventorySummary>;
};

function Dashboard(props: SharedProps & { isSample: boolean; setPage: (page: Page) => void }) {
  const [metric, setMetric] = useState<"revenue" | "profit">("revenue");
  const points = useMemo(() => trend(props.filtered, props.period, props.range), [props.filtered, props.period, props.range]);
  const priorPoints = useMemo(() => trend(props.previous, props.period, previousRange(props.range)), [props.previous, props.period, props.range]);
  const [selectedKey, setSelectedKey] = useState("");
  const [selectedProduct, setSelectedProduct] = useState("");
  const selected = points.find((point) => point.key === selectedKey) || points[points.length - 1];
  const current = metrics(props.filtered), prior = metrics(props.previous);
  const lowStock = props.inventory.filter((item) => item.status === "Low Stock" || item.status === "Out of Stock");
  const best = productSummary(props.filtered).slice(0, 5);
  const categories = categorySummary(props.filtered);
  const revenueChange = comparison(current.revenue, prior.revenue);
  const insights = [
    revenueChange !== null ? `Sales are ${Math.abs(revenueChange).toFixed(1)}% ${revenueChange >= 0 ? "higher" : "lower"} than the previous period.` : null,
    best[0] ? `${best[0].product} is the best-selling item with ${count(best[0].units)} units sold.` : null,
    categories[0] ? `${categories[0].category} brought in the most sales at ${money(categories[0].revenue)}.` : null,
    lowStock[0]?.daysRemaining !== null && lowStock[0]?.daysRemaining !== undefined ? `${lowStock[0].product} may run out in about ${Math.ceil(lowStock[0].daysRemaining)} days at the current sales rate.` : null,
  ].filter(Boolean) as string[];
  const selectedIndex = selected ? points.findIndex((point) => point.key === selected.key) : -1;
  const previousForPoint = selectedIndex >= 0 ? priorPoints[selectedIndex]?.sales || [] : [];
  return <div>
    <PageHeading title="Namaste! Here’s your shop summary" text="See sales, popular items and stock that needs attention." />
    {props.isSample && <div className="sample-note"><Icon name="alert" /><span><b>You are viewing demo data.</b> Upload your own sales file to see your shop’s results.</span><Button kind="plain" onClick={() => props.setPage("upload")}>Upload my data <Icon name="arrow" size={15} /></Button></div>}
    <PeriodSelector {...props} />
    <div className="kpi-grid">
      <Kpi label="Total Sales" value={money(current.revenue)} icon="rupee" change={comparison(current.revenue, prior.revenue)} />
      <Kpi label="Estimated Profit" value={current.profit === null ? "Not available" : money(current.profit)} icon="report" change={current.profit !== null && prior.profit !== null ? comparison(current.profit, prior.profit) : null} />
      <Kpi label="Items Sold" value={count(current.units)} icon="cart" change={comparison(current.units, prior.units)} />
      <Kpi label="Low-Stock Products" value={props.inventory.some((item) => item.stock !== undefined) ? String(lowStock.length) : "Not available"} icon="box" onClick={() => props.setPage("products")} />
    </div>
    <Card className="chart-card"><SalesChart points={points} metric={metric} setMetric={setMetric} hasProfit={current.hasCost} selected={selected} setSelected={(point) => setSelectedKey(point.key)} />{selected && <DetailPanel point={selected} previous={previousForPoint} />}</Card>
    <div className="dashboard-bottom">
      <Card className="simple-list"><SectionTitle title="Best-Selling Items" text="Click a bar to see its sales value" />{best.length ? <><HorizontalBarChart items={best.map((item) => ({ label: item.product, value: item.units, note: `${count(item.units)} sold` }))} selected={selectedProduct} onSelect={setSelectedProduct} />{selectedProduct && <div className="chart-selection"><b>{selectedProduct}</b><span>{money(best.find((item) => item.product === selectedProduct)?.revenue || 0)} total sales in this period</span></div>}</> : <Empty title="No items found" text="There are no sales in this period." />}</Card>
      <Card className="simple-list"><SectionTitle title="Products Running Out" text="Based on current stock and sales" action={<button onClick={() => props.setPage("products")}>View Products <Icon name="arrow" size={14} /></button>} />{lowStock.slice(0, 4).map((item) => <div className="stock-row" key={item.product}><div><b>{item.product}</b><small>{item.stock} in stock · {item.averageDaily.toFixed(1)} sold/day</small></div><span>{item.daysRemaining === null ? "Days unknown" : `~${Math.ceil(item.daysRemaining)} days left`}</span></div>)}{!lowStock.length && <Empty title="No low-stock items" text="All products with stock data are above their reorder level." />}</Card>
      <Card className="simple-list insights"><SectionTitle title="Important Insights" text="Calculated for the selected period" />{insights.length ? insights.slice(0, 4).map((text, index) => <button key={text} onClick={() => index === 3 ? props.setPage("products") : undefined}><span>{index + 1}</span><p>{text}</p><Icon name="arrow" size={15} /></button>) : <Empty title="Not enough data" text="Choose a longer period to generate useful insights." />}</Card>
    </div>
  </div>;
}

function SectionTitle({ title, text, action }: { title: string; text: string; action?: ReactNode }) {
  return <div className="section-title"><div><b>{title}</b><span>{text}</span></div>{action}</div>;
}

function Upload({ rawRows, setRawRows, mapping, setMapping, dataset, setDataset, setIsSample, onAnalysed }: {
  rawRows: RawRow[]; setRawRows: (rows: RawRow[]) => void; mapping: Mapping; setMapping: (mapping: Mapping) => void;
  dataset: string; setDataset: (name: string) => void; setIsSample: (value: boolean) => void; onAnalysed: () => void;
}) {
  const input = useRef<HTMLInputElement>(null);
  const [error, setError] = useState("");
  const headers = Object.keys(rawRows[0] || {});
  const report = quality(rawRows, mapping);
  async function readFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]; if (!file) return;
    setError("");
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
      const rows = XLSX.utils.sheet_to_json<RawRow>(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
      if (!rows.length) throw new Error("The file does not contain any data rows.");
      setRawRows(rows); setMapping(detectMapping(Object.keys(rows[0]))); setDataset(file.name); setIsSample(false);
    } catch (reason) { setError(reason instanceof Error ? reason.message : "We could not read this file."); }
  }
  const canAnalyse = report.unmappedCore.length === 0 && report.valid > 0;
  return <div><PageHeading title="Upload Sales Data" text="Upload your CSV or Excel file. We’ll check it before making your dashboard." />
    <input ref={input} type="file" accept=".csv,.xlsx,.xls" hidden onChange={readFile} />
    <Card className="upload-box"><div className="upload-icon"><Icon name="upload" size={28} /></div><b>Choose your sales file</b><p>CSV, XLS or XLSX · Up to 25 MB</p><Button onClick={() => input.current?.click()}>Browse Files</Button>{error && <div className="error-message"><Icon name="alert" />{error}</div>}</Card>
    <div className="upload-summary"><Card><span>File</span><b>{dataset}</b><small>{rawRows.length.toLocaleString("en-IN")} rows found</small></Card><Card><span>Data quality</span><b>{report.valid} valid rows</b><small>{report.invalid} invalid · {report.duplicates} duplicate</small></Card><Card><span>Ready?</span><b>{canAnalyse ? "Yes" : "Needs attention"}</b><small>{report.unmappedCore.length ? `Map ${report.unmappedCore.map((key) => FIELD_LABELS[key]).join(", ")}` : "Required columns are mapped"}</small></Card></div>
    <Card className="mapping"><SectionTitle title="Match Your Columns" text="Check that we understood your file correctly" /><div className="mapping-list">{(Object.keys(FIELD_LABELS) as (keyof Mapping)[]).map((key) => <label key={key}><span>{FIELD_LABELS[key]}{["date", "product", "quantity", "price"].includes(key) && <i>Required</i>}</span><select value={mapping[key]} onChange={(event) => setMapping({ ...mapping, [key]: event.target.value })}><option value="">Not available</option>{headers.map((header) => <option key={header}>{header}</option>)}</select></label>)}</div></Card>
    <Card className="preview"><SectionTitle title="File Preview" text={`Showing first ${Math.min(6, rawRows.length)} rows`} /><div className="table-scroll"><table><thead><tr>{headers.slice(0, 8).map((header) => <th key={header}>{header}</th>)}</tr></thead><tbody>{rawRows.slice(0, 6).map((row, index) => <tr key={index}>{headers.slice(0, 8).map((header) => <td key={header}>{String(row[header])}</td>)}</tr>)}</tbody></table></div><div className="analyse-row"><span>{canAnalyse ? "Your file is ready." : "Please map the required columns above."}</span>{canAnalyse && <Button onClick={onAnalysed}>Analyze My Data <Icon name="arrow" size={15} /></Button>}</div></Card>
  </div>;
}

function Reports(props: SharedProps & { dataset: string }) {
  const data = metrics(props.filtered), prior = metrics(props.previous);
  const products = productSummary(props.filtered), categories = categorySummary(props.filtered);
  const [sort, setSort] = useState<"date" | "sales">("date");
  const [selectedProduct, setSelectedProduct] = useState("");
  const [selectedCategory, setSelectedCategory] = useState("");
  const [selectedDay, setSelectedDay] = useState("");
  const rows = trend(props.filtered, props.period, props.range).sort((a, b) => sort === "sales" ? b.revenue - a.revenue : a.key.localeCompare(b.key));
  function exportCsv() {
    const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(props.filtered.map((sale) => ({ Date: formatISO(sale.date), Product: sale.product, Category: sale.category, Quantity: sale.quantity, Price: sale.price, Cost: sale.cost ?? "", Transaction: sale.transaction ?? "" }))));
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    const link = document.createElement("a"); link.href = url; link.download = `${props.dataset}-${props.period}-report.csv`; link.click(); URL.revokeObjectURL(url);
  }
  return <div><PageHeading title="Sales Reports" text="One simple report for any time period." action={<Button kind="secondary" icon="download" onClick={exportCsv}>Export CSV</Button>} /><PeriodSelector {...props} /><div className="report-kpis"><Kpi label="Total Sales" value={money(data.revenue)} icon="rupee" change={comparison(data.revenue, prior.revenue)} /><Kpi label="Estimated Profit" value={data.profit === null ? "Not available" : money(data.profit)} icon="report" /><Kpi label="Items Sold" value={count(data.units)} icon="cart" /><Kpi label="Transactions" value={data.transactions ? String(data.transactions) : "Not available"} icon="people" /></div>
    <div className="visual-grid">
      <Card className="visual-card"><SectionTitle title="Best-Selling Products" text="Interactive horizontal bar chart" /><HorizontalBarChart items={products.slice(0, 7).map((item) => ({ label: item.product, value: item.units, note: `${count(item.units)} units` }))} selected={selectedProduct} onSelect={setSelectedProduct} />{selectedProduct && <div className="chart-selection"><b>{selectedProduct}</b><span>{money(products.find((item) => item.product === selectedProduct)?.revenue || 0)} sales</span></div>}</Card>
      <Card className="visual-card"><SectionTitle title="Sales by Category" text="Click a category to inspect its contribution" />{categories.length ? <CategoryDonut items={categories} selected={selectedCategory} onSelect={setSelectedCategory} /> : <Empty title="No category data" text="There are no sales in this period." />}{selectedCategory && <div className="chart-selection"><b>{selectedCategory}</b><span>{money(categories.find((item) => item.category === selectedCategory)?.revenue || 0)} sales in this period</span></div>}</Card>
      <Card className="visual-card weekday-card"><SectionTitle title="Sales by Day of Week" text="Click a bar to view the exact sales amount" /><WeekdayChart sales={props.filtered} selected={selectedDay} onSelect={setSelectedDay} />{selectedDay && <div className="chart-selection"><b>{selectedDay}</b><span>{money(props.filtered.filter((sale) => sale.date.toLocaleDateString("en-IN", { weekday: "long" }) === selectedDay).reduce((sum, sale) => sum + sale.quantity * sale.price, 0))} total sales</span></div>}</Card>
    </div>
    <Card className="report-table"><SectionTitle title="Sales Trend Table" text={`${rows.length} periods`} action={<select value={sort} onChange={(event) => setSort(event.target.value as "date" | "sales")}><option value="date">Sort by date</option><option value="sales">Sort by sales</option></select>} /><div className="table-scroll"><table><thead><tr><th>Period</th><th>Sales</th><th>Profit</th><th>Items Sold</th><th>Transactions</th><th>Best Seller</th></tr></thead><tbody>{rows.map((row) => { const rowMetrics = metrics(row.sales); return <tr key={row.key}><td>{row.label}</td><td>{money(row.revenue)}</td><td>{row.profit === null ? "Not available" : money(row.profit)}</td><td>{count(rowMetrics.units)}</td><td>{rowMetrics.transactions || "—"}</td><td>{productSummary(row.sales)[0]?.product || "—"}</td></tr>; })}</tbody></table></div></Card>
  </div>;
}

function Products({ allSales, filtered, inventory, overrides, setOverrides }: { allSales: Sale[]; filtered: Sale[]; inventory: ReturnType<typeof inventorySummary>; overrides: Record<string, InventoryOverride>; setOverrides: (value: Record<string, InventoryOverride>) => void }) {
  const [search, setSearch] = useState(""), [category, setCategory] = useState("All"), [sort, setSort] = useState("stock");
  const [editing, setEditing] = useState<string | null>(null);
  const [adding, setAdding] = useState(false);
  const categories = ["All", ...new Set(inventory.map((item) => item.category))];
  const filteredSummary = productSummary(filtered);
  const rows = inventory.map((item) => ({ ...item, periodUnits: filteredSummary.find((product) => product.product === item.product)?.units || 0 })).filter((item) => item.product.toLowerCase().includes(search.toLowerCase()) && (category === "All" || item.category === category)).sort((a, b) => sort === "sales" ? b.periodUnits - a.periodUnits : (a.stock ?? Infinity) - (b.stock ?? Infinity));
  function save(product: string, stock: number, reorder: number) { setOverrides({ ...overrides, [product]: { ...overrides[product], stock, reorder } }); setEditing(null); }
  return <div><PageHeading title="Products & Stock" text="Search products, update stock and see what needs reordering." action={<Button kind="secondary" icon="plus" onClick={() => setAdding(true)}>Add Product</Button>} />
    <div className="product-tools"><label className="search"><Icon name="search" /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search product" /></label><select value={category} onChange={(event) => setCategory(event.target.value)}>{categories.map((item) => <option key={item}>{item}</option>)}</select><select value={sort} onChange={(event) => setSort(event.target.value)}><option value="stock">Lowest stock first</option><option value="sales">Highest sales first</option></select></div>
    {!allSales.some((sale) => sale.stock !== undefined) && !Object.keys(overrides).length && <div className="warning-note"><Icon name="alert" /><span>Your sales file has no current stock column. Sales results still work, and you can enter stock manually by selecting a product.</span></div>}
    <Card className="product-table"><div className="table-scroll"><table><thead><tr><th>Product Name</th><th>Category</th><th>Stock Available</th><th>Units Sold</th><th>Selling Price</th><th>Purchase Cost</th><th>Reorder Level</th><th>Status</th></tr></thead><tbody>{rows.map((item) => <tr key={item.product} onClick={() => setEditing(item.product)}><td><b>{item.product}</b></td><td>{item.category}</td><td>{item.stock ?? "Not entered"}</td><td>{count(item.periodUnits)}</td><td>{money(item.price)}</td><td>{item.costPrice === undefined ? "Not available" : money(item.costPrice)}</td><td>{item.reorder}</td><td><Status value={item.status} /></td></tr>)}</tbody></table></div></Card>
    {editing && <ProductEditor item={inventory.find((item) => item.product === editing)!} close={() => setEditing(null)} save={save} />}
    {adding && <AddProduct close={() => setAdding(false)} add={(product, details) => { setOverrides({ ...overrides, [product]: details }); setAdding(false); }} />}
  </div>;
}

function AddProduct({ close, add }: { close: () => void; add: (product: string, details: InventoryOverride) => void }) {
  const [product, setProduct] = useState(""), [category, setCategory] = useState("Other");
  const [stock, setStock] = useState(0), [reorder, setReorder] = useState(10), [price, setPrice] = useState(0), [cost, setCost] = useState(0);
  return <div className="drawer-backdrop" onClick={close}><div onClick={(event) => event.stopPropagation()}><Card className="drawer"><button className="drawer-close" onClick={close}><Icon name="close" /></button><div className="drawer-title">Add a New Product</div><p>This adds an inventory record only. It does not create a past sale.</p><label className="edit-field">Product name<input value={product} onChange={(event) => setProduct(event.target.value)} /></label><label className="edit-field">Category<input value={category} onChange={(event) => setCategory(event.target.value)} /></label><label className="edit-field">Stock available<input type="number" min="0" value={stock} onChange={(event) => setStock(Number(event.target.value))} /></label><label className="edit-field">Reorder level<input type="number" min="0" value={reorder} onChange={(event) => setReorder(Number(event.target.value))} /></label><label className="edit-field">Selling price<input type="number" min="0" value={price} onChange={(event) => setPrice(Number(event.target.value))} /></label><label className="edit-field">Purchase cost<input type="number" min="0" value={cost} onChange={(event) => setCost(Number(event.target.value))} /></label><Button onClick={() => product.trim() && add(product.trim(), { category: category.trim() || "Other", stock, reorder, price, cost })}>Add Product</Button></Card></div></div>;
}

function ProductEditor({ item, close, save }: { item: ReturnType<typeof inventorySummary>[number]; close: () => void; save: (product: string, stock: number, reorder: number) => void }) {
  const [stock, setStock] = useState(item.stock ?? 0), [reorder, setReorder] = useState(item.reorder);
  return <div className="drawer-backdrop" onClick={close}><div onClick={(event) => event.stopPropagation()}><Card className="drawer"><button className="drawer-close" onClick={close}><Icon name="close" /></button><div className="drawer-title">{item.product}</div><p>{item.category} · {count(item.units)} units in sales history</p><div className="drawer-stats"><div><span>Average daily sales</span><b>{item.averageDaily.toFixed(1)}</b></div><div><span>Days remaining</span><b>{item.daysRemaining === null ? "Unknown" : Math.ceil(item.daysRemaining)}</b></div><div><span>Suggested reorder</span><b>{item.suggested || "Not needed"}</b></div></div><label className="edit-field">Stock available<input type="number" min="0" value={stock} onChange={(event) => setStock(Number(event.target.value))} /></label><label className="edit-field">Reorder level<input type="number" min="0" value={reorder} onChange={(event) => setReorder(Number(event.target.value))} /></label><Button onClick={() => save(item.product, stock, reorder)}>Save Stock Details</Button></Card></div></div>;
}

function Status({ value }: { value: string }) { return <span className={`stock-status ${value.toLowerCase().replaceAll(" ", "-")}`}>{value}</span>; }

function Customers({ filtered, previous, rangeLabel }: { filtered: Sale[]; previous: Sale[]; rangeLabel: string }) {
  const pairs = basketPairs(filtered), previousPairs = basketPairs(previous);
  const [selected, setSelected] = useState<(typeof pairs)[number] | null>(null);
  const hasTransactions = filtered.some((sale) => sale.transaction);
  return <div><PageHeading title="Customer Purchase Insights" text="See which products are commonly bought in the same bill." />
    {!hasTransactions ? <Empty title="Transaction IDs are needed" text="Your file must include Transaction ID, Date and Product Name to find products bought together reliably." /> : <><div className="basket-summary"><Card><span>Bills analysed</span><b>{new Set(filtered.map((sale) => sale.transaction).filter(Boolean)).size}</b></Card><Card><span>Product pairs found</span><b>{pairs.length}</b></Card><Card><span>Selected period</span><b>{rangeLabel}</b></Card></div><div className="pair-grid">{pairs.slice(0, 9).map((pair) => <Card className="pair-card" key={`${pair.a}-${pair.b}`} onClick={() => setSelected(pair)}><div className="pair-icon"><Icon name="cart" /></div><div><b>{pair.a}</b><span>often bought with</span><b>{pair.b}</b></div><div className="pair-numbers"><strong>{pair.count} shared bills</strong><span>{pair.percentage}% of {pair.a} bills</span></div><Icon name="arrow" /></Card>)}</div>{!pairs.length && <Empty title="No product combinations found" text="Bills in this period contain only one unique product, or the period has no sales." />}</>}
    {selected && <div className="drawer-backdrop" onClick={() => setSelected(null)}><div onClick={(event) => event.stopPropagation()}><Card className="drawer pair-detail"><button className="drawer-close" onClick={() => setSelected(null)}><Icon name="close" /></button><div className="drawer-title">{selected.a} + {selected.b}</div><p>Purchase combination details for {rangeLabel}</p><div className="drawer-stats"><div><span>Shared transactions</span><b>{selected.count}</b></div><div><span>Sales from these bills</span><b>{money(selected.revenue)}</b></div><div><span>Previous period</span><b>{previousPairs.find((pair) => pair.a === selected.a && pair.b === selected.b)?.count ?? "Not enough data"}</b></div></div><div className="plain-note">In {selected.percentage}% of bills containing {selected.a}, customers also bought {selected.b}.</div></Card></div></div>}
  </div>;
}

function Future({ allSales, inventory }: { allSales: Sale[]; inventory: ReturnType<typeof inventorySummary> }) {
  const [horizon, setHorizon] = useState(7);
  const result = allSales.length ? forecast(allSales, horizon) : { points: [], average: 0, sufficient: false, accuracy: null };
  const historyRange = allSales.length ? { start: new Date(Math.min(...allSales.map((sale) => sale.date.getTime()))), end: new Date(Math.max(...allSales.map((sale) => sale.date.getTime()))), label: "" } : getRange("month");
  const history = trend(allSales, "custom", historyRange).slice(-30);
  const restock = inventory.filter((item) => item.daysRemaining !== null && item.daysRemaining <= horizon);
  return <div><PageHeading title="Future Sales Estimates" text="A simple estimate based on your shop’s recent daily sales." /><div className="future-options"><span>Estimate for</span>{[[7, "Next 7 days"], [30, "Next 30 days"], [90, "Next 3 months"]].map(([days, label]) => <button key={days} className={horizon === days ? "active" : ""} onClick={() => setHorizon(days as number)}>{label}</button>)}</div>
    {!result.sufficient ? <Empty title="More sales history is needed" text="At least 14 days of sales are required for a useful moving-average estimate. We will not make up future sales." /> : <><div className="forecast-kpis"><Kpi label="Estimated Future Sales" value={money(result.average * horizon)} icon="future" /><Kpi label="Average Per Day" value={money(result.average)} icon="rupee" /><Kpi label="May Need Restocking" value={String(restock.length)} icon="box" /></div><Card className="chart-card"><SalesChart points={history} metric="revenue" setMetric={() => {}} hasProfit={false} futurePoints={result.points.map((point) => ({ label: point.date.toLocaleDateString("en-IN", { day: "numeric", month: "short" }), value: point.value }))} /><div className="forecast-legend"><span><i />Actual sales</span><span><i />Estimated sales</span></div></Card><div className="forecast-bottom"><Card className="simple-list"><SectionTitle title="Products That May Need Restocking" text={`Within the next ${horizon} days`} />{restock.slice(0, 6).map((item) => <div className="stock-row" key={item.product}><div><b>{item.product}</b><small>{item.stock} in stock · {item.averageDaily.toFixed(1)} sold/day</small></div><span>Order about {item.suggested || Math.ceil(item.averageDaily * horizon)} units</span></div>)}{!restock.length && <Empty title="No immediate restocking found" text="Current stock appears sufficient for this estimate." />}</Card><Card className="assumptions"><SectionTitle title="How This Estimate Works" text="Simple and transparent" /><ul><li>Uses the average daily sales from recent history.</li><li>Assumes upcoming demand is similar to recent demand.</li><li>Does not account for festivals, weather or supplier changes.</li><li>{result.accuracy === null ? "Accuracy cannot yet be evaluated." : `Recent holdout WAPE: ${result.accuracy.toFixed(1)}% (lower is better).`}</li></ul></Card></div></>}
  </div>;
}

export default App;
