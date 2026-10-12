export type RawRow = Record<string, string | number>;
export type Period = "today" | "week" | "month" | "year" | "custom";

export type Mapping = {
  date: string;
  product: string;
  category: string;
  quantity: string;
  price: string;
  cost: string;
  stock: string;
  reorder: string;
  transaction: string;
  customer: string;
};

export type Sale = {
  date: Date;
  product: string;
  category: string;
  quantity: number;
  price: number;
  cost?: number;
  stock?: number;
  reorder?: number;
  transaction?: string;
  customer?: string;
};

export type DateRange = { start: Date; end: Date; label: string };
export type InventoryOverride = { stock: number; reorder: number; category?: string; price?: number; cost?: number };

export const FIELD_LABELS: Record<keyof Mapping, string> = {
  date: "Date",
  product: "Product Name",
  category: "Category",
  quantity: "Quantity Sold",
  price: "Selling Price",
  cost: "Purchase Cost",
  stock: "Current Stock",
  reorder: "Reorder Level",
  transaction: "Transaction ID",
  customer: "Customer ID",
};

const aliases: Record<keyof Mapping, string[]> = {
  date: ["date", "sale date", "transaction date", "datetime", "timestamp"],
  product: ["product", "product name", "item", "item name", "description"],
  category: ["category", "product category", "group"],
  quantity: ["quantity", "quantity sold", "qty", "units", "units sold"],
  price: ["price", "selling price", "sales price", "rate", "unit price"],
  cost: ["cost", "purchase cost", "buying price", "unit cost"],
  stock: ["stock", "current stock", "stock quantity", "stock available", "inventory"],
  reorder: ["reorder", "reorder level", "minimum stock", "min stock"],
  transaction: ["transaction", "transaction id", "invoice", "invoice id", "bill no", "bill number"],
  customer: ["customer", "customer id", "buyer id"],
};

export function detectMapping(headers: string[]): Mapping {
  return Object.fromEntries(
    (Object.keys(aliases) as (keyof Mapping)[]).map((key) => [
      key,
      headers.find((header) => aliases[key].includes(header.trim().toLowerCase())) || "",
    ]),
  ) as Mapping;
}

export function normalizeRows(rows: RawRow[], mapping: Mapping): Sale[] {
  const get = (row: RawRow, key: keyof Mapping) => mapping[key] ? row[mapping[key]] : "";
  return rows.flatMap((row) => {
    const date = parseDate(get(row, "date"));
    const product = String(get(row, "product") ?? "").trim();
    const quantity = Number(get(row, "quantity"));
    const price = Number(get(row, "price"));
    if (!date || !product || !Number.isFinite(quantity) || quantity < 0 || !Number.isFinite(price) || price < 0) return [];
    const optionalNumber = (key: keyof Mapping) => {
      const raw = get(row, key);
      if (raw === "" || raw == null) return undefined;
      const value = Number(raw);
      return Number.isFinite(value) && value >= 0 ? value : undefined;
    };
    return [{
      date,
      product,
      category: String(get(row, "category") || "Other").trim(),
      quantity,
      price,
      cost: optionalNumber("cost"),
      stock: optionalNumber("stock"),
      reorder: optionalNumber("reorder"),
      transaction: String(get(row, "transaction") || "").trim() || undefined,
      customer: String(get(row, "customer") || "").trim() || undefined,
    }];
  });
}

function parseDate(value: string | number): Date | null {
  if (typeof value === "number" && value > 20000) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(excelEpoch.getTime() + value * 86400000);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const text = String(value || "").trim();
  if (!text) return null;
  const indian = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(.*)$/);
  const date = indian
    ? new Date(`${indian[3]}-${indian[2].padStart(2, "0")}-${indian[1].padStart(2, "0")}${indian[4]}`)
    : new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

const startOfDay = (date: Date) => new Date(date.getFullYear(), date.getMonth(), date.getDate());
const endOfDay = (date: Date) => new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
const addDays = (date: Date, days: number) => new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);

export function getRange(period: Period, customStart?: string, customEnd?: string, anchor = new Date()): DateRange {
  const today = startOfDay(anchor);
  if (period === "today") return { start: today, end: endOfDay(today), label: "Today" };
  if (period === "week") {
    const monday = addDays(today, -((today.getDay() + 6) % 7));
    return { start: monday, end: endOfDay(addDays(monday, 6)), label: "This Week" };
  }
  if (period === "month") return { start: new Date(today.getFullYear(), today.getMonth(), 1), end: endOfDay(new Date(today.getFullYear(), today.getMonth() + 1, 0)), label: "This Month" };
  if (period === "year") return { start: new Date(today.getFullYear(), 0, 1), end: endOfDay(new Date(today.getFullYear(), 11, 31)), label: "This Year" };
  const start = customStart ? startOfDay(new Date(`${customStart}T00:00:00`)) : addDays(today, -29);
  const end = customEnd ? endOfDay(new Date(`${customEnd}T00:00:00`)) : endOfDay(today);
  return { start, end, label: "Custom Range" };
}

export function previousRange(range: DateRange): DateRange {
  const duration = range.end.getTime() - range.start.getTime() + 1;
  const end = new Date(range.start.getTime() - 1);
  return { start: new Date(end.getTime() - duration + 1), end, label: "Previous Period" };
}

export function filterSales(sales: Sale[], range: DateRange) {
  return sales.filter((sale) => sale.date >= range.start && sale.date <= range.end);
}

export function metrics(sales: Sale[]) {
  const revenue = sales.reduce((sum, sale) => sum + sale.quantity * sale.price, 0);
  const hasCost = sales.length > 0 && sales.every((sale) => sale.cost !== undefined);
  const profit = hasCost ? sales.reduce((sum, sale) => sum + sale.quantity * (sale.price - (sale.cost || 0)), 0) : null;
  const units = sales.reduce((sum, sale) => sum + sale.quantity, 0);
  const transactions = new Set(sales.map((sale) => sale.transaction).filter(Boolean)).size;
  return { revenue, profit, units, transactions, hasCost };
}

export function productSummary(sales: Sale[]) {
  const map = new Map<string, { product: string; category: string; units: number; revenue: number; cost: number | null; price: number }>();
  sales.forEach((sale) => {
    const item = map.get(sale.product) || { product: sale.product, category: sale.category, units: 0, revenue: 0, cost: sale.cost === undefined ? null : 0, price: sale.price };
    item.units += sale.quantity;
    item.revenue += sale.quantity * sale.price;
    item.price = sale.price;
    if (item.cost !== null && sale.cost !== undefined) item.cost += sale.quantity * sale.cost;
    else item.cost = null;
    map.set(sale.product, item);
  });
  return [...map.values()].sort((a, b) => b.units - a.units);
}

export function categorySummary(sales: Sale[]) {
  const map = new Map<string, number>();
  sales.forEach((sale) => map.set(sale.category, (map.get(sale.category) || 0) + sale.quantity * sale.price));
  return [...map].map(([category, revenue]) => ({ category, revenue })).sort((a, b) => b.revenue - a.revenue);
}

export type TrendPoint = { key: string; label: string; revenue: number; profit: number | null; sales: Sale[] };

export function trend(sales: Sale[], period: Period, range: DateRange): TrendPoint[] {
  const days = Math.ceil((range.end.getTime() - range.start.getTime()) / 86400000);
  const mode = period === "today" && sales.some((sale) => sale.date.getHours() !== 0) ? "hour" : period === "year" || days > 90 ? "month" : days > 45 ? "week" : "day";
  const buckets = new Map<string, Sale[]>();
  sales.forEach((sale) => {
    const date = sale.date;
    let key: string;
    if (mode === "hour") key = `${String(date.getHours()).padStart(2, "0")}:00`;
    else if (mode === "month") key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    else if (mode === "week") {
      const monday = addDays(startOfDay(date), -((date.getDay() + 6) % 7));
      key = formatISO(monday);
    } else key = formatISO(date);
    buckets.set(key, [...(buckets.get(key) || []), sale]);
  });
  return [...buckets].sort(([a], [b]) => a.localeCompare(b)).map(([key, items]) => {
    const total = metrics(items);
    const label = mode === "hour" ? key : mode === "month" ? new Date(`${key}-01T00:00:00`).toLocaleDateString("en-IN", { month: "short" }) : new Date(`${key}T00:00:00`).toLocaleDateString("en-IN", { day: "numeric", month: "short" });
    return { key, label, revenue: total.revenue, profit: total.profit, sales: items };
  });
}

export function basketPairs(sales: Sale[]) {
  if (!sales.some((sale) => sale.transaction)) return [];
  const transactions = new Map<string, Sale[]>();
  sales.forEach((sale) => sale.transaction && transactions.set(sale.transaction, [...(transactions.get(sale.transaction) || []), sale]));
  const productTransactions = new Map<string, Set<string>>();
  const pairs = new Map<string, { a: string; b: string; count: number; revenue: number }>();
  transactions.forEach((items, transaction) => {
    const products = [...new Set(items.map((item) => item.product))].sort();
    products.forEach((product) => productTransactions.set(product, new Set([...(productTransactions.get(product) || []), transaction])));
    for (let i = 0; i < products.length; i++) for (let j = i + 1; j < products.length; j++) {
      const key = `${products[i]}|||${products[j]}`;
      const pair = pairs.get(key) || { a: products[i], b: products[j], count: 0, revenue: 0 };
      pair.count += 1;
      pair.revenue += items.reduce((sum, item) => sum + item.quantity * item.price, 0);
      pairs.set(key, pair);
    }
  });
  return [...pairs.values()].map((pair) => ({ ...pair, percentage: Math.round(pair.count / (productTransactions.get(pair.a)?.size || 1) * 100) })).sort((a, b) => b.count - a.count);
}

export function inventorySummary(allSales: Sale[], overrides: Record<string, InventoryOverride>) {
  const uniqueDays = Math.max(new Set(allSales.map((sale) => formatISO(sale.date))).size, 1);
  const products = productSummary(allSales);
  const salesProducts = new Set(products.map((product) => product.product));
  const manualProducts = Object.entries(overrides).filter(([product]) => !salesProducts.has(product)).map(([product, details]) => ({
    product, category: details.category || "Other", units: 0, revenue: 0, cost: details.cost === undefined ? null : 0, price: details.price || 0,
  }));
  return [...products, ...manualProducts].map((product) => {
    const source = [...allSales].reverse().find((sale) => sale.product === product.product);
    const stock = overrides[product.product]?.stock ?? source?.stock;
    const reorder = overrides[product.product]?.reorder ?? source?.reorder ?? 10;
    const averageDaily = product.units / uniqueDays;
    const daysRemaining = stock !== undefined && averageDaily > 0 && uniqueDays >= 2 ? stock / averageDaily : null;
    const suggested = stock !== undefined && stock < reorder ? Math.max(Math.ceil(averageDaily * 14 - stock), reorder - stock) : 0;
    const status = stock === undefined ? "Unknown" : stock === 0 ? "Out of Stock" : stock < reorder ? "Low Stock" : "In Stock";
    return { ...product, stock, reorder, averageDaily, daysRemaining, suggested, status, costPrice: source?.cost ?? overrides[product.product]?.cost };
  });
}

export function quality(rows: RawRow[], mapping: Mapping) {
  const serialized = rows.map((row) => JSON.stringify(row));
  const duplicates = serialized.length - new Set(serialized).size;
  const valid = normalizeRows(rows, mapping);
  return { total: rows.length, valid: valid.length, invalid: rows.length - valid.length, duplicates, unmappedCore: (["date", "product", "quantity", "price"] as (keyof Mapping)[]).filter((key) => !mapping[key]) };
}

export function forecast(sales: Sale[], horizon: number) {
  const dailyMap = new Map<string, number>();
  sales.forEach((sale) => dailyMap.set(formatISO(sale.date), (dailyMap.get(formatISO(sale.date)) || 0) + sale.quantity * sale.price));
  const daily = [...dailyMap].sort(([a], [b]) => a.localeCompare(b)).map(([date, revenue]) => ({ date, revenue }));
  if (daily.length < 14) return { points: [], average: 0, sufficient: false, accuracy: null };
  const recent = daily.slice(-Math.min(28, daily.length));
  const average = recent.reduce((sum, point) => sum + point.revenue, 0) / recent.length;
  const lastDate = sales.reduce((a, b) => a.date > b.date ? a : b).date;
  const points = Array.from({ length: horizon }, (_, index) => ({ date: addDays(lastDate, index + 1), value: average }));
  const holdout = daily.slice(-7);
  const training = daily.slice(-35, -7);
  const baseline = training.reduce((sum, point) => sum + point.revenue, 0) / training.length;
  const errors = holdout.map((point) => Math.abs(point.revenue - baseline));
  const wape = holdout.reduce((sum, point) => sum + point.revenue, 0) ? errors.reduce((a, b) => a + b, 0) / holdout.reduce((sum, point) => sum + point.revenue, 0) * 100 : null;
  return { points, average, sufficient: true, accuracy: wape };
}

export function comparison(current: number, previous: number) {
  if (!previous) return null;
  return ((current - previous) / previous) * 100;
}

export function formatISO(date: Date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

export function makeSampleRows(): RawRow[] {
  const products = [
    ["Amul Milk 1L", "Dairy", 72, 66, 18, 12],
    ["Parle-G Biscuits", "Snacks", 10, 8, 42, 20],
    ["Tata Salt 1kg", "Staples", 28, 24, 35, 12],
    ["Aashirvaad Atta 5kg", "Staples", 286, 252, 14, 10],
    ["Fortune Oil 1L", "Cooking", 146, 131, 8, 10],
    ["Red Label Tea 500g", "Beverages", 275, 241, 16, 8],
    ["Britannia Bread", "Bakery", 48, 39, 6, 8],
    ["Surf Excel 1kg", "Home Care", 142, 126, 22, 8],
  ] as const;
  const rows: RawRow[] = [];
  const today = startOfDay(new Date());
  for (let day = 119; day >= 0; day--) {
    const date = addDays(today, -day);
    for (let transaction = 0; transaction < 4; transaction++) {
      const tx = `S-${formatISO(date)}-${transaction + 1}`;
      const first = products[(day + transaction) % products.length];
      const second = products[(day + transaction + (transaction % 2 ? 1 : 5)) % products.length];
      [first, second].forEach((product, item) => rows.push({
        Date: formatISO(date),
        "Product Name": product[0],
        Category: product[1],
        "Quantity Sold": 1 + ((day + transaction + item) % 4),
        "Selling Price": product[2],
        "Purchase Cost": product[3],
        "Current Stock": product[4],
        "Reorder Level": product[5],
        "Transaction ID": tx,
      }));
    }
  }
  return rows;
}
