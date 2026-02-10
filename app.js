/* global docx */
const DB_NAME = "kaidan-pwa";
const DB_VERSION = 1;
const DEFAULT_COMPANY_TITLE = "毕节共利食品有限责任公司-销货单";
const DEFAULT_ACCOUNT_TEXT = "刘正彬 6215582406000752975 中国工商银行宜宾市翠屏区西郊支行\n刘正彬 6228482469624921172 中国农业银行宜宾市翠屏区西郊支行";
const DEFAULT_AI_PROMPT = "你是结构化信息抽取助手，只输出严格 JSON，不要输出多余文字。JSON字段为：{\"customer\":\"\", \"destination\":\"\", \"plate\":\"\", \"driver\":\"\", \"date\":\"\", \"items\":[{\"product\":\"\", \"spec_jin\":\"\", \"count\":\"\", \"price_per_ton\":\"\"}]} 图片通常包含多行产品，每行一条 items。 重要：spec_jin/count/price_per_ton 只输出纯数字（不要单位、不要斜杠、不要中文）。 规则：spec_jin 来自“(22斤/件)”中的 22；count 来自“小计/共计325件”中的 325； price_per_ton 来自“价格15030/吨”中的 15030。 示例行：红毛7只(22斤/件): ... 小计325件, 价格15030/吨。 看不清的字段填空字符串 。";
const CATEGORIES = ["自动", "鸡", "鸡副", "混合"];

const state = {
  db: null,
  customers: [],
  products: [],
  vehicles: [],
  invoices: [],
  invoiceItems: [],
  invoiceAudit: [],
  statView: "query",
  statRegionFilterProv: null,
};

const qs = (sel) => document.querySelector(sel);
const qsa = (sel) => Array.from(document.querySelectorAll(sel));

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (event) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains("settings")) {
        db.createObjectStore("settings", { keyPath: "key" });
      }
      if (!db.objectStoreNames.contains("customers")) {
        const store = db.createObjectStore("customers", { keyPath: "id", autoIncrement: true });
        store.createIndex("name", "name", { unique: true });
      }
      if (!db.objectStoreNames.contains("products")) {
        const store = db.createObjectStore("products", { keyPath: "id", autoIncrement: true });
        store.createIndex("name", "name", { unique: true });
      }
      if (!db.objectStoreNames.contains("vehicles")) {
        const store = db.createObjectStore("vehicles", { keyPath: "id", autoIncrement: true });
        store.createIndex("plate1", "plate1", { unique: true });
      }
      if (!db.objectStoreNames.contains("invoices")) {
        const store = db.createObjectStore("invoices", { keyPath: "id", autoIncrement: true });
        store.createIndex("date_iso", "date_iso", { unique: false });
      }
      if (!db.objectStoreNames.contains("invoice_items")) {
        const store = db.createObjectStore("invoice_items", { keyPath: "id", autoIncrement: true });
        store.createIndex("invoice_id", "invoice_id", { unique: false });
      }
      if (!db.objectStoreNames.contains("invoice_audit")) {
        const store = db.createObjectStore("invoice_audit", { keyPath: "id", autoIncrement: true });
        store.createIndex("invoice_id", "invoice_id", { unique: false });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function tx(storeNames, mode = "readonly") {
  return state.db.transaction(storeNames, mode);
}

function getAll(storeName) {
  return new Promise((resolve, reject) => {
    const t = tx([storeName]);
    const store = t.objectStore(storeName);
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

function getByKey(storeName, key) {
  return new Promise((resolve, reject) => {
    const t = tx([storeName]);
    const store = t.objectStore(storeName);
    const req = store.get(key);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

function putItem(storeName, value) {
  return new Promise((resolve, reject) => {
    const t = tx([storeName], "readwrite");
    const store = t.objectStore(storeName);
    const req = store.put(value);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function addItem(storeName, value) {
  return new Promise((resolve, reject) => {
    const t = tx([storeName], "readwrite");
    const store = t.objectStore(storeName);
    const req = store.add(value);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function deleteItem(storeName, key) {
  return new Promise((resolve, reject) => {
    const t = tx([storeName], "readwrite");
    const store = t.objectStore(storeName);
    const req = store.delete(key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

function clearStore(storeName) {
  return new Promise((resolve, reject) => {
    const t = tx([storeName], "readwrite");
    const store = t.objectStore(storeName);
    const req = store.clear();
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

async function getSetting(key, fallback = "") {
  const row = await getByKey("settings", key);
  return row ? row.value : fallback;
}

async function setSetting(key, value) {
  await putItem("settings", { key, value });
}

function toFloat(val, fallback = 0) {
  const n = parseFloat(String(val || "").replace(/,/g, ""));
  return Number.isFinite(n) ? n : fallback;
}

function toInt(val, fallback = 0) {
  const n = parseInt(String(val || ""), 10);
  return Number.isFinite(n) ? n : fallback;
}

function formatDecimal(num) {
  if (!Number.isFinite(num)) return "";
  let s = num.toFixed(6);
  s = s.replace(/0+$/, "").replace(/\.$/, "");
  return s;
}

function formatAmountCn(value) {
  const v = Math.round(toFloat(value));
  if (!Number.isFinite(v)) return String(value || "");
  if (v >= 10000) {
    const wan = Math.floor(v / 10000);
    const rest = v % 10000;
    return rest ? `${wan}万${rest}元` : `${wan}万元`;
  }
  return `${v}元`;
}

function normalizeText(s) {
  if (!s) return "";
  return String(s).replace(/\s+/g, "").replace(/[：:，,。.!？?（）()【】\[\]{}<>《》-]/g, "");
}

function parseDateToIso(s) {
  try {
    const parts = String(s || "").trim().split(/[./-]/);
    if (parts.length >= 3) {
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      const d = parseInt(parts[2], 10);
      if (Number.isFinite(y) && Number.isFinite(m) && Number.isFinite(d)) {
        const dt = new Date(y, m - 1, d);
        if (!Number.isNaN(dt.getTime())) {
          return dt.toISOString().slice(0, 10);
        }
      }
    }
  } catch (err) {
    return "";
  }
  return "";
}

function formatDateInput(value) {
  const parts = String(value || "").trim().split(/[./-]/);
  if (parts.length >= 3) {
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const d = parseInt(parts[2], 10);
    if (Number.isFinite(y) && Number.isFinite(m) && Number.isFinite(d)) {
      return `${y}.${m}.${d}`;
    }
  }
  return value;
}

function safeFilename(text) {
  return String(text || "")
    .replace(/[<>:\\/|?*]/g, "")
    .replace(/[\n\r]/g, " ")
    .trim();
}

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}.${d.getMonth() + 1}.${d.getDate()}`;
}

function monthStartEnd(year, month) {
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 0);
  return [start, end];
}

function quarterStartEnd(year, quarter) {
  const m = (quarter - 1) * 3 + 1;
  return monthStartEnd(year, m).map((d, idx) => (idx === 0 ? d : new Date(year, m + 2, 0)));
}

function addMonths(dateObj, months) {
  const d = new Date(dateObj.getTime());
  const day = d.getDate();
  d.setDate(1);
  d.setMonth(d.getMonth() + months);
  const last = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  d.setDate(Math.min(day, last));
  return d;
}

function resolveProvince(location) {
  if (!location) return "未知";
  const loc = String(location).trim();
  const provinces = [
    "北京","天津","上海","重庆","河北","山西","辽宁","吉林","黑龙江","江苏","浙江","安徽","福建","江西",
    "山东","河南","湖北","湖南","广东","海南","四川","贵州","云南","陕西","甘肃","青海","内蒙古","广西","西藏",
    "宁夏","新疆"
  ];
  for (const prov of provinces) {
    if (loc.includes(prov)) return prov;
  }
  const loc2 = loc.replace(/(市|省|自治区|自治州|地区|盟|州)$/g, "");
  if (CITY_TO_PROVINCE[loc2]) return CITY_TO_PROVINCE[loc2];
  for (const [city, prov] of Object.entries(CITY_TO_PROVINCE)) {
    if (loc.includes(city)) return prov;
  }
  return "未知";
}

const CITY_TO_PROVINCE = {
  "北京": "北京","北京市": "北京","天津": "天津","天津市": "天津","上海": "上海","上海市": "上海","重庆": "重庆","重庆市": "重庆",
  "石家庄": "河北","唐山": "河北","秦皇岛": "河北","邯郸": "河北","邢台": "河北","保定": "河北","张家口": "河北","承德": "河北",
  "沧州": "河北","廊坊": "河北","衡水": "河北","太原": "山西","大同": "山西","朔州": "山西","忻州": "山西","阳泉": "山西",
  "吕梁": "山西","晋中": "山西","长治": "山西","晋城": "山西","临汾": "山西","运城": "山西","沈阳": "辽宁","大连": "辽宁",
  "鞍山": "辽宁","抚顺": "辽宁","本溪": "辽宁","丹东": "辽宁","锦州": "辽宁","营口": "辽宁","阜新": "辽宁","辽阳": "辽宁",
  "盘锦": "辽宁","铁岭": "辽宁","朝阳": "辽宁","葫芦岛": "辽宁","长春": "吉林","吉林": "吉林","四平": "吉林","辽源": "吉林",
  "通化": "吉林","白山": "吉林","松原": "吉林","白城": "吉林","延边朝鲜族自治州": "吉林","哈尔滨": "黑龙江",
  "齐齐哈尔": "黑龙江","牡丹江": "黑龙江","佳木斯": "黑龙江","大庆": "黑龙江","鸡西": "黑龙江","鹤岗": "黑龙江",
  "双鸭山": "黑龙江","伊春": "黑龙江","七台河": "黑龙江","黑河": "黑龙江","绥化": "黑龙江","大兴安岭地区": "黑龙江",
  "南京": "江苏","无锡": "江苏","徐州": "江苏","常州": "江苏","苏州": "江苏","南通": "江苏","连云港": "江苏","淮安": "江苏",
  "盐城": "江苏","扬州": "江苏","镇江": "江苏","泰州": "江苏","宿迁": "江苏","杭州": "浙江","宁波": "浙江","温州": "浙江",
  "嘉兴": "浙江","湖州": "浙江","绍兴": "浙江","金华": "浙江","衢州": "浙江","舟山": "浙江","台州": "浙江","丽水": "浙江",
  "合肥": "安徽","芜湖": "安徽","蚌埠": "安徽","淮南": "安徽","马鞍山": "安徽","淮北": "安徽","铜陵": "安徽",
  "安庆": "安徽","黄山": "安徽","滁州": "安徽","阜阳": "安徽","宿州": "安徽","六安": "安徽","亳州": "安徽",
  "池州": "安徽","宣城": "安徽","福州": "福建","厦门": "福建","莆田": "福建","三明": "福建","泉州": "福建","漳州": "福建",
  "南平": "福建","龙岩": "福建","宁德": "福建","南昌": "江西","景德镇": "江西","萍乡": "江西","九江": "江西","新余": "江西",
  "鹰潭": "江西","赣州": "江西","吉安": "江西","宜春": "江西","抚州": "江西","上饶": "江西","济南": "山东","青岛": "山东",
  "淄博": "山东","枣庄": "山东","东营": "山东","烟台": "山东","潍坊": "山东","济宁": "山东","泰安": "山东","威海": "山东",
  "日照": "山东","临沂": "山东","德州": "山东","聊城": "山东","滨州": "山东","菏泽": "山东","郑州": "河南","开封": "河南",
  "洛阳": "河南","平顶山": "河南","安阳": "河南","鹤壁": "河南","新乡": "河南","焦作": "河南","濮阳": "河南","许昌": "河南",
  "漯河": "河南","三门峡": "河南","南阳": "河南","商丘": "河南","信阳": "河南","周口": "河南","驻马店": "河南","武汉": "湖北",
  "黄石": "湖北","十堰": "湖北","宜昌": "湖北","襄阳": "湖北","鄂州": "湖北","荆门": "湖北","孝感": "湖北","荆州": "湖北",
  "黄冈": "湖北","咸宁": "湖北","随州": "湖北","恩施土家族苗族自治州": "湖北","仙桃": "湖北","潜江": "湖北","天门": "湖北",
  "神农架林区": "湖北","长沙": "湖南","株洲": "湖南","湘潭": "湖南","衡阳": "湖南","邵阳": "湖南","岳阳": "湖南",
  "常德": "湖南","张家界": "湖南","益阳": "湖南","郴州": "湖南","永州": "湖南","怀化": "湖南","娄底": "湖南",
  "湘西土家族苗族自治州": "湖南","广州": "广东","韶关": "广东","深圳": "广东","珠海": "广东","汕头": "广东","佛山": "广东",
  "江门": "广东","湛江": "广东","茂名": "广东","肇庆": "广东","惠州": "广东","梅州": "广东","汕尾": "广东","河源": "广东",
  "阳江": "广东","清远": "广东","东莞": "广东","中山": "广东","潮州": "广东","揭阳": "广东","云浮": "广东","海口": "海南",
  "三亚": "海南","三沙": "海南","儋州": "海南","成都": "四川","自贡": "四川","攀枝花": "四川","泸州": "四川","德阳": "四川",
  "绵阳": "四川","广元": "四川","遂宁": "四川","内江": "四川","乐山": "四川","南充": "四川","眉山": "四川","宜宾": "四川",
  "广安": "四川","达州": "四川","雅安": "四川","巴中": "四川","资阳": "四川","阿坝藏族羌族自治州": "四川","甘孜藏族自治州": "四川",
  "凉山彝族自治州": "四川","贵阳": "贵州","六盘水": "贵州","遵义": "贵州","安顺": "贵州","毕节": "贵州","铜仁": "贵州",
  "黔西南布依族苗族自治州": "贵州","黔东南苗族侗族自治州": "贵州","黔南布依族苗族自治州": "贵州","昆明": "云南","曲靖": "云南",
  "玉溪": "云南","保山": "云南","昭通": "云南","丽江": "云南","普洱": "云南","临沧": "云南","楚雄彝族自治州": "云南",
  "红河哈尼族彝族自治州": "云南","文山壮族苗族自治州": "云南","西双版纳傣族自治州": "云南","大理白族自治州": "云南",
  "德宏傣族景颇族自治州": "云南","怒江傈僳族自治州": "云南","迪庆藏族自治州": "云南","西安": "陕西","铜川": "陕西","宝鸡": "陕西",
  "咸阳": "陕西","渭南": "陕西","延安": "陕西","汉中": "陕西","榆林": "陕西","安康": "陕西","商洛": "陕西","兰州": "甘肃",
  "嘉峪关": "甘肃","金昌": "甘肃","白银": "甘肃","天水": "甘肃","武威": "甘肃","张掖": "甘肃","平凉": "甘肃","酒泉": "甘肃",
  "庆阳": "甘肃","定西": "甘肃","陇南": "甘肃","临夏回族自治州": "甘肃","甘南藏族自治州": "甘肃","西宁": "青海","海东": "青海",
  "海北藏族自治州": "青海","黄南藏族自治州": "青海","海南藏族自治州": "青海","果洛藏族自治州": "青海","玉树藏族自治州": "青海",
  "海西蒙古族藏族自治州": "青海","呼和浩特": "内蒙古","包头": "内蒙古","乌海": "内蒙古","赤峰": "内蒙古","通辽": "内蒙古",
  "鄂尔多斯": "内蒙古","呼伦贝尔": "内蒙古","巴彦淖尔": "内蒙古","乌兰察布": "内蒙古","兴安盟": "内蒙古","锡林郭勒盟": "内蒙古",
  "阿拉善盟": "内蒙古","南宁": "广西","柳州": "广西","桂林": "广西","梧州": "广西","北海": "广西","防城港": "广西",
  "钦州": "广西","贵港": "广西","玉林": "广西","百色": "广西","贺州": "广西","河池": "广西","来宾": "广西","崇左": "广西",
  "拉萨": "西藏","日喀则": "西藏","昌都": "西藏","林芝": "西藏","山南": "西藏","那曲": "西藏","阿里地区": "西藏",
  "银川": "宁夏","石嘴山": "宁夏","吴忠": "宁夏","固原": "宁夏","中卫": "宁夏","乌鲁木齐": "新疆","克拉玛依": "新疆",
  "吐鲁番": "新疆","哈密": "新疆","昌吉回族自治州": "新疆","博尔塔拉蒙古自治州": "新疆","巴音郭楞蒙古自治州": "新疆",
  "阿克苏地区": "新疆","克孜勒苏柯尔克孜自治州": "新疆","喀什地区": "新疆","和田地区": "新疆","伊犁哈萨克自治州": "新疆",
  "塔城地区": "新疆","阿勒泰地区": "新疆","石河子": "新疆","阿拉尔": "新疆","图木舒克": "新疆","五家渠": "新疆","北屯": "新疆",
  "铁门关": "新疆","双河": "新疆","可克达拉": "新疆","昆玉": "新疆","胡杨河": "新疆"
};

function setStatus(text) {
  qs("#appStatus").textContent = text;
}

function nextTick() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

async function commitActiveInput() {
  const active = document.activeElement;
  if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA")) {
    active.blur();
    await nextTick();
  }
}

function initTabs() {
  qsa(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      qsa(".tab-btn").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const tab = btn.dataset.tab;
      qsa(".tab-panel").forEach((panel) => panel.classList.remove("active"));
      qs(`#tab-${tab}`).classList.add("active");
    });
  });
}

function bindInvoiceInputs() {
  const itemInputs = qsa("#itemsTable input");
  itemInputs.forEach((input) => {
    input.addEventListener("input", () => {
      const row = input.dataset.row;
      const field = input.dataset.field;
      if (field === "weight" || field === "amount" || field === "price") {
        if (input.value.trim()) {
          input.dataset.manual = "1";
          delete input.dataset.auto;
        } else {
          delete input.dataset.manual;
          delete input.dataset.auto;
        }
      }
      if (field === "spec" || field === "count") {
        const weightInput = getItemInput(row, "weight");
        const amountInput = getItemInput(row, "amount");
        if (weightInput && !weightInput.dataset.manual) {
          weightInput.dataset.auto = "1";
        }
        if (amountInput && !amountInput.dataset.manual) {
          amountInput.dataset.auto = "1";
        }
      }
      if (field === "price") {
        const amountInput = getItemInput(row, "amount");
        if (amountInput && !amountInput.dataset.manual) {
          amountInput.dataset.auto = "1";
        }
      }
      if (input.dataset.field === "name") {
        autoFillProduct(row);
      }
      recalcRow(row);
      recalcTotals();
    });
    input.addEventListener("blur", () => {
      const row = input.dataset.row;
      recalcRow(row);
      recalcTotals();
    });
  });

  const customerInput = qs("#invCustomer");
  const driverInput = qs("#invDriver");
  const phoneInput = qs("#invPhone");
  const plate1Input = qs("#invPlate1");
  const plate2Input = qs("#invPlate2");

  customerInput.addEventListener("input", fillFromCustomer);
  customerInput.addEventListener("blur", fillFromCustomer);
  customerInput.addEventListener("change", fillFromCustomer);

  [driverInput, phoneInput, plate1Input, plate2Input].forEach((el) => {
    el.addEventListener("input", fillVehicleByAny);
    el.addEventListener("blur", fillVehicleByAny);
    el.addEventListener("change", fillVehicleByAny);
  });
}

function autoFillProduct(row) {
  const name = getItemValue(row, "name");
  if (!name) return;
  const match = state.products.find((p) => p.name === name);
  if (!match) return;
  const specInput = getItemInput(row, "spec");
  if (specInput && !specInput.value.trim() && match.spec_jin !== null && match.spec_jin !== undefined) {
    specInput.value = formatDecimal(toFloat(match.spec_jin));
  }
}

function getItemInput(row, field) {
  return qs(`#itemsTable input[data-row="${row}"][data-field="${field}"]`);
}

function getItemValue(row, field) {
  const input = getItemInput(row, field);
  return input ? input.value.trim() : "";
}

function recalcRow(row) {
  const specInput = getItemInput(row, "spec");
  const countInput = getItemInput(row, "count");
  const weightInput = getItemInput(row, "weight");
  const priceInput = getItemInput(row, "price");
  const amountInput = getItemInput(row, "amount");

  const spec = toFloat(specInput.value, NaN);
  const count = toFloat(countInput.value, NaN);
  const weightIsAuto = weightInput.dataset.auto === "1" && !weightInput.dataset.manual;
  const amountIsAuto = amountInput.dataset.auto === "1" && !amountInput.dataset.manual;
  const priceIsAuto = priceInput.dataset.auto === "1" && !priceInput.dataset.manual;

  const weight = weightIsAuto ? NaN : toFloat(weightInput.value, NaN);
  const price = priceIsAuto ? NaN : toFloat(priceInput.value, NaN);
  const amount = amountIsAuto ? NaN : toFloat(amountInput.value, NaN);

  let calcWeight = Number.isFinite(weight) ? weight : NaN;
  if (!Number.isFinite(calcWeight) && Number.isFinite(spec) && Number.isFinite(count)) {
    calcWeight = (spec * count) / 2000;
  }

  let calcAmount = Number.isFinite(amount) ? amount : NaN;
  if (!Number.isFinite(calcAmount) && Number.isFinite(price) && Number.isFinite(calcWeight)) {
    calcAmount = price * calcWeight;
  }

  let calcPrice = Number.isFinite(price) ? price : NaN;
  if (!Number.isFinite(calcPrice) && Number.isFinite(calcAmount) && Number.isFinite(calcWeight) && calcWeight !== 0) {
    calcPrice = calcAmount / calcWeight;
  }

  const canAutoWeight = !weightInput.value.trim() || weightInput.dataset.auto === "1";
  const canAutoPrice = !priceInput.value.trim() || priceInput.dataset.auto === "1";
  const canAutoAmount = !amountInput.value.trim() || amountInput.dataset.auto === "1";

  if (canAutoWeight && Number.isFinite(calcWeight)) {
    weightInput.value = calcWeight.toFixed(3);
    weightInput.dataset.auto = "1";
  }
  if (canAutoPrice && Number.isFinite(calcPrice)) {
    priceInput.value = formatDecimal(calcPrice);
    priceInput.dataset.auto = "1";
  }
  if (canAutoAmount && Number.isFinite(calcAmount)) {
    amountInput.value = Math.round(calcAmount).toString();
    amountInput.dataset.auto = "1";
  }
}

function recalcTotals() {
  let totalQty = 0;
  let totalWeight = 0;
  let totalAmount = 0;
  for (let i = 0; i < 5; i += 1) {
    totalQty += toFloat(getItemValue(i, "count"));
    totalWeight += toFloat(getItemValue(i, "weight"));
    totalAmount += toFloat(getItemValue(i, "amount"));
  }
  qs("#totalsLine").textContent = `总件数：${totalQty.toFixed(0)}    总吨位：${totalWeight.toFixed(3)}    总价格：${totalAmount.toFixed(0)}`;
  const summary = qs("#invSummary");
  if (!summary.value.trim()) {
    summary.value = `共${totalQty.toFixed(0)}件，${totalWeight.toFixed(3)}吨，合计${totalAmount.toFixed(2)}元`;
  }
}

function gatherItems() {
  const items = [];
  for (let i = 0; i < 5; i += 1) {
    const name = getItemValue(i, "name");
    const spec = getItemValue(i, "spec");
    const count = getItemValue(i, "count");
    const weight = getItemValue(i, "weight");
    const price = getItemValue(i, "price");
    const amount = getItemValue(i, "amount");
    if (name || spec || count || weight || price || amount) {
      items.push({
        name,
        spec,
        count,
        weight,
        price,
        amount,
      });
    }
  }
  return items;
}

function resolveCategory(items) {
  const cat = qs("#invCategory").value.trim();
  if (cat && cat !== "自动") return cat;
  let hasJifu = false;
  let hasJi = false;
  let matched = 0;
  let total = 0;
  items.forEach((it) => {
    const name = it.name.trim();
    if (!name) return;
    total += 1;
    const match = state.products.find((p) => p.name === name);
    if (match) {
      const c = (match.category || "").trim();
      if (c === "鸡副") hasJifu = true;
      if (c === "鸡") hasJi = true;
      matched += 1;
    }
  });
  if (hasJifu && hasJi) return "混合";
  if (total > 0 && matched === total) {
    if (hasJifu && !hasJi) return "鸡副";
    if (hasJi && !hasJifu) return "鸡";
  }
  if (hasJifu) return "鸡副";
  if (hasJi) return "鸡";
  return "鸡";
}

function fillFromCustomer() {
  const name = qs("#invCustomer").value.trim();
  if (!name) return;
  const nkey = normalizeText(name);
  let match = state.customers.find((c) => normalizeText(c.name) === nkey);
  if (!match) {
    match = state.customers.find((c) => normalizeText(c.name).includes(nkey));
  }
  if (!match) return;
  if (!qs("#invLocation").value.trim() && match.location) {
    qs("#invLocation").value = match.location;
  }
}

function normalizeValue(val) {
  if (!val) return "";
  return String(val).replace(/\s+/g, "").trim();
}

function findVehicleMatch(value, field) {
  if (!value) return null;
  const v = normalizeValue(value);
  if (!v) return null;
  const exact = state.vehicles.find((row) => normalizeValue(row[field] || "") === v);
  if (exact) return exact;
  return state.vehicles.find((row) => normalizeValue(row[field] || "").includes(v));
}

function fillVehicleByAny() {
  const plate1 = qs("#invPlate1").value.trim();
  const plate2 = qs("#invPlate2").value.trim();
  const driver = qs("#invDriver").value.trim();
  const phone = qs("#invPhone").value.trim();

  let match =
    findVehicleMatch(plate1, "plate1") ||
    findVehicleMatch(plate2, "plate2") ||
    findVehicleMatch(driver, "driver") ||
    findVehicleMatch(phone, "phone");

  if (!match) return;

  if (!plate1 && match.plate1) qs("#invPlate1").value = match.plate1;
  if (!plate2 && match.plate2) qs("#invPlate2").value = match.plate2;
  if (!driver && match.driver) qs("#invDriver").value = match.driver;
  if (!phone && match.phone) qs("#invPhone").value = match.phone;
}

async function refreshReferenceData() {
  state.customers = await getAll("customers");
  state.products = await getAll("products");
  state.vehicles = await getAll("vehicles");

  const customerList = qs("#customerList");
  customerList.innerHTML = state.customers.map((c) => `<option value="${c.name}"></option>`).join("");

  const productList = qs("#productList");
  productList.innerHTML = state.products.map((p) => `<option value="${p.name}"></option>`).join("");

  const driverList = qs("#driverList");
  driverList.innerHTML = state.vehicles.map((v) => `<option value="${v.driver || ""}"></option>`).join("");

  const phoneList = qs("#phoneList");
  phoneList.innerHTML = state.vehicles.map((v) => `<option value="${v.phone || ""}"></option>`).join("");

  const plateList1 = qs("#plateList1");
  if (plateList1) {
    plateList1.innerHTML = state.vehicles.map((v) => `<option value="${v.plate1 || ""}"></option>`).join("");
  }
  const plateList2 = qs("#plateList2");
  if (plateList2) {
    plateList2.innerHTML = state.vehicles.map((v) => `<option value="${v.plate2 || ""}"></option>`).join("");
  }

  const locationList = qs("#locationList");
  const locations = new Set(state.customers.map((c) => c.location).filter(Boolean));
  locationList.innerHTML = Array.from(locations).map((l) => `<option value="${l}"></option>`).join("");
}

function renderTable(tableId, rows, columns) {
  const table = qs(tableId);
  const tbody = table.querySelector("tbody");
  tbody.innerHTML = rows
    .map((row) =>
      `<tr data-id="${row.id}">${columns.map((col) => `<td>${row[col] ?? ""}</td>`).join("")}</tr>`
    )
    .join("");
}

async function renderCustomers() {
  state.customers = await getAll("customers");
  renderTable("#cusTable", state.customers, ["name", "location", "province", "note"]);
}

async function renderProducts() {
  state.products = await getAll("products");
  renderTable("#proTable", state.products, ["name", "spec_jin", "category", "note"]);
}

async function renderVehicles() {
  state.vehicles = await getAll("vehicles");
  renderTable("#vehTable", state.vehicles, ["plate1", "plate2", "driver", "phone"]);
}

function bindTableSelection(tableId, handler) {
  const table = qs(tableId);
  table.addEventListener("click", (e) => {
    const row = e.target.closest("tr");
    if (!row || !row.dataset.id) return;
    handler(row.dataset.id);
  });
}

function bindMultiSelection(tableId) {
  let lastIndex = null;
  const table = qs(tableId);
  table.addEventListener("click", (e) => {
    const tr = e.target.closest("tr");
    if (!tr) return;
    const rows = qsa(`${tableId} tbody tr`);
    const idx = rows.indexOf(tr);
    if (e.shiftKey && lastIndex !== null) {
      const [start, end] = idx > lastIndex ? [lastIndex, idx] : [idx, lastIndex];
      rows.forEach((r, i) => {
        if (i >= start && i <= end) {
          r.classList.add("selected");
        }
      });
    } else if (e.metaKey || e.ctrlKey) {
      tr.classList.toggle("selected");
      lastIndex = idx;
      return;
    } else {
      rows.forEach((r) => r.classList.remove("selected"));
      tr.classList.add("selected");
    }
    lastIndex = idx;
  });
}

function getSelectedIds(tableId) {
  return qsa(`${tableId} tbody tr.selected`).map((r) => Number(r.dataset.id));
}

async function saveCustomer() {
  const name = qs("#cusName").value.trim();
  if (!name) return alert("请输入客户名称");
  const location = qs("#cusLocation").value.trim();
  const province = qs("#cusProvince").value.trim() || resolveProvince(location);
  const payload = {
    name,
    location,
    province,
    note: qs("#cusNote").value.trim(),
  };
  const existing = state.customers.find((c) => c.name === name);
  if (existing) {
    payload.id = existing.id;
  }
  await putItem("customers", payload);
  await renderCustomers();
  await refreshReferenceData();
}

async function deleteCustomer() {
  const ids = getSelectedIds("#cusTable");
  if (!ids.length) return alert("请选择要删除的客户");
  if (!confirm(`确定删除选中的 ${ids.length} 条客户？`)) return;
  for (const id of ids) {
    await deleteItem("customers", id);
  }
  await renderCustomers();
  await refreshReferenceData();
}

async function saveProduct() {
  const name = qs("#proName").value.trim();
  if (!name) return alert("请输入产品名称");
  const payload = {
    name,
    spec_jin: qs("#proSpec").value.trim() ? toFloat(qs("#proSpec").value) : "",
    category: qs("#proCategory").value.trim(),
    note: qs("#proNote").value.trim(),
  };
  const existing = state.products.find((p) => p.name === name);
  if (existing) {
    payload.id = existing.id;
  }
  await putItem("products", payload);
  await renderProducts();
  await refreshReferenceData();
}

async function deleteProduct() {
  const ids = getSelectedIds("#proTable");
  if (!ids.length) return alert("请选择要删除的产品");
  if (!confirm(`确定删除选中的 ${ids.length} 条产品？`)) return;
  for (const id of ids) {
    await deleteItem("products", id);
  }
  await renderProducts();
  await refreshReferenceData();
}

async function saveVehicle() {
  const plate1 = qs("#vehPlate1").value.trim();
  if (!plate1) return alert("请输入车牌号1");
  const payload = {
    plate1,
    plate2: qs("#vehPlate2").value.trim(),
    driver: qs("#vehDriver").value.trim(),
    phone: qs("#vehPhone").value.trim(),
  };
  const existing = state.vehicles.find((v) => v.plate1 === plate1);
  if (existing) {
    payload.id = existing.id;
  }
  await putItem("vehicles", payload);
  await renderVehicles();
  await refreshReferenceData();
}

async function deleteVehicle() {
  const ids = getSelectedIds("#vehTable");
  if (!ids.length) return alert("请选择要删除的车辆");
  if (!confirm(`确定删除选中的 ${ids.length} 条车辆？`)) return;
  for (const id of ids) {
    await deleteItem("vehicles", id);
  }
  await renderVehicles();
  await refreshReferenceData();
}

function clearForm(prefix) {
  qsa(`${prefix} input`).forEach((el) => {
    if (el.type !== "file") el.value = "";
  });
  qsa(`${prefix} textarea`).forEach((el) => {
    el.value = "";
  });
}

async function ensureDefaults() {
  const defaults = {
    company_title: DEFAULT_COMPANY_TITLE,
    account_text: DEFAULT_ACCOUNT_TEXT,
    last_invoice_no: "0",
    ai_provider: "Gemini",
    ai_proxy_url: "",
    ai_api_url: "",
    ai_model: "gemini-2.0-flash",
    ai_api_key: "",
    ai_prompt: DEFAULT_AI_PROMPT,
  };
  for (const [key, value] of Object.entries(defaults)) {
    const existing = await getSetting(key, null);
    if (existing === null) {
      await setSetting(key, value);
    }
  }
}

async function loadSettings() {
  qs("#setCompanyTitle").value = await getSetting("company_title", DEFAULT_COMPANY_TITLE);
  qs("#setAccountText").value = await getSetting("account_text", DEFAULT_ACCOUNT_TEXT);
  qs("#setAiProvider").value = await getSetting("ai_provider", "Gemini");
  qs("#setAiProxyUrl").value = await getSetting("ai_proxy_url", "");
  qs("#setAiUrl").value = await getSetting("ai_api_url", "");
  qs("#setAiModel").value = await getSetting("ai_model", "gemini-2.0-flash");
  qs("#setAiKey").value = await getSetting("ai_api_key", "");
  qs("#setAiPrompt").value = await getSetting("ai_prompt", DEFAULT_AI_PROMPT);
  const lastNo = toInt(await getSetting("last_invoice_no", "0"), 0);
  qs("#invNo").value = String(lastNo + 1);
  if (!qs("#invDate").value.trim()) {
    qs("#invDate").value = todayStr();
  }
}

async function saveSettings() {
  await setSetting("company_title", qs("#setCompanyTitle").value.trim());
  await setSetting("account_text", qs("#setAccountText").value.trim());
  alert("设置已保存");
}

async function saveAiSettings() {
  await setSetting("ai_provider", qs("#setAiProvider").value.trim());
  await setSetting("ai_proxy_url", qs("#setAiProxyUrl").value.trim());
  await setSetting("ai_api_url", qs("#setAiUrl").value.trim());
  await setSetting("ai_model", qs("#setAiModel").value.trim());
  await setSetting("ai_api_key", qs("#setAiKey").value.trim());
  await setSetting("ai_prompt", qs("#setAiPrompt").value.trim());
  alert("AI 设置已保存");
}

function extractJson(text) {
  if (!text) return null;
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try {
    return JSON.parse(match[0]);
  } catch (err) {
    return null;
  }
}

function getAiStatusEl() {
  return qs("#aiStatus");
}

function setAiStatus(text) {
  const el = getAiStatusEl();
  if (el) el.textContent = text;
}

function setAiDuration(text) {
  const el = qs("#aiDuration");
  if (el) el.textContent = text;
}

function setAiTestStatus(text) {
  const el = qs("#aiTestStatus");
  if (el) el.textContent = text;
}

async function imageToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function extractAiContent(data) {
  if (!data) return "";
  if (data.choices && data.choices[0] && data.choices[0].message) {
    return data.choices[0].message.content || "";
  }
  if (data.candidates && data.candidates[0]) {
    const parts = data.candidates[0].content && data.candidates[0].content.parts;
    if (Array.isArray(parts)) {
      return parts.map((p) => p.text || "").join("");
    }
  }
  return "";
}

async function aiRecognizeImage(file) {
  const proxyUrl = (await getSetting("ai_proxy_url", "")).trim();
  const apiUrl = (await getSetting("ai_api_url", "")).trim();
  const apiKey = (await getSetting("ai_api_key", "")).trim();
  const model = (await getSetting("ai_model", "")).trim();
  const prompt = (await getSetting("ai_prompt", DEFAULT_AI_PROMPT)).trim();

  if (!apiUrl && !proxyUrl) throw new Error("请先设置 AI API 地址或代理地址");
  if (!apiKey && !proxyUrl) throw new Error("请先设置 AI API Key");
  if (!model) throw new Error("请先设置模型名称");

  const dataUrl = await imageToDataUrl(file);

  const isProxy = !!proxyUrl;
  const body = isProxy
    ? {
        apiUrl,
        apiKey,
        model,
        prompt,
        imageDataUrl: dataUrl,
      }
    : {
        model,
        messages: [
          { role: "system", content: prompt },
          {
            role: "user",
            content: [
              { type: "text", text: "请从图片中提取并输出 JSON" },
              { type: "image_url", image_url: { url: dataUrl } },
            ],
          },
        ],
        temperature: 0,
      };

  const resp = await fetch(isProxy ? proxyUrl : apiUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      ...(isProxy ? {} : { Authorization: `Bearer ${apiKey}` }),
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`AI 接口错误: ${resp.status} ${text}`);
  }

  const data = await resp.json();
  const content = extractAiContent(data);
  const json = extractJson(content);
  if (!json) throw new Error("AI 返回内容无法解析为 JSON");
  return json;
}

async function testAiApi() {
  const proxyUrl = (await getSetting("ai_proxy_url", "")).trim();
  const apiUrl = (await getSetting("ai_api_url", "")).trim();
  const apiKey = (await getSetting("ai_api_key", "")).trim();
  const model = (await getSetting("ai_model", "")).trim();
  const prompt = (await getSetting("ai_prompt", DEFAULT_AI_PROMPT)).trim();

  if (!apiUrl && !proxyUrl) return alert("请先设置 AI API 地址或代理地址");
  if (!apiKey && !proxyUrl) return alert("请先设置 AI API Key");
  if (!model) return alert("请先设置模型名称");

  if (location.protocol !== "https:" && location.hostname !== "localhost" && location.hostname !== "127.0.0.1") {
    setAiTestStatus("当前页面非 HTTPS，iOS 可能阻止请求");
  }

  setAiTestStatus("测试中…");
  try {
    const isProxy = !!proxyUrl;
    const body = isProxy
      ? {
          apiUrl,
          apiKey,
          model,
          prompt,
          imageDataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO3w2O0AAAAASUVORK5CYII=",
        }
      : {
          model,
          messages: [
            { role: "system", content: prompt },
            { role: "user", content: "只返回空 JSON：{}" },
          ],
          temperature: 0,
        };

    const resp = await fetch(isProxy ? proxyUrl : apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(isProxy ? {} : { Authorization: `Bearer ${apiKey}` }),
      },
      body: JSON.stringify(body),
    });

    const text = await resp.text();
    if (!resp.ok) {
      setAiTestStatus(`失败：${resp.status}`);
      alert(`接口返回错误：${resp.status}\n${text.slice(0, 300)}`);
      return;
    }
    setAiTestStatus("成功：接口可访问");
  } catch (err) {
    setAiTestStatus("失败：Failed to fetch");
    alert(
      "请求失败（Failed to fetch）。常见原因：\n" +
        "1. API 不支持浏览器跨域（CORS）\n" +
        "2. 证书或 HTTPS 问题\n" +
        "3. 网络被系统拦截\n\n" +
        `错误详情：${err.message || err}`
    );
  }
}

async function fillInvoiceFromAi(result) {
  if (!result) return;
  const customer = (result.customer || "").trim();
  const destination = (result.destination || "").trim();
  const plate = (result.plate || "").trim();
  const driver = (result.driver || "").trim();
  const date = (result.date || "").trim();

  if (customer) qs("#invCustomer").value = customer;
  if (destination) qs("#invLocation").value = destination;
  if (plate) qs("#invPlate1").value = plate;
  if (driver) qs("#invDriver").value = driver;
  if (date) qs("#invDate").value = formatDateInput(date);

  // 强制拉最新数据库，保证 AI 填入后可正确自动关联
  await refreshReferenceData();

  // 自动补全：客户 -> 到达地点；车牌/司机/电话 -> 互补字段
  fillFromCustomer();
  fillVehicleByAny();
  await nextTick();
  fillFromCustomer();
  fillVehicleByAny();

  const items = Array.isArray(result.items) ? result.items : [];
  for (let i = 0; i < 5; i += 1) {
    const it = items[i] || {};
    const name = (it.product || "").trim();
    const spec = (it.spec_jin || "").trim();
    const count = (it.count || "").trim();
    const price = (it.price_per_ton || "").trim();
    if (name) getItemInput(i, "name").value = name;
    if (spec) getItemInput(i, "spec").value = spec;
    if (count) getItemInput(i, "count").value = count;
    if (price) getItemInput(i, "price").value = price;
    recalcRow(i);
  }
  recalcTotals();
}

async function aiImportAndExport(file) {
  const start = Date.now();
  setAiStatus("识别中…");
  const result = await aiRecognizeImage(file);
  await fillInvoiceFromAi(result);
  await exportInvoice();
  setAiStatus("识别完成并已导出");
  const secs = (Date.now() - start) / 1000;
  setAiDuration(`用时 ${secs.toFixed(1)} 秒`);
}

function getTotals() {
  let totalQty = 0;
  let totalWeight = 0;
  let totalAmount = 0;
  for (let i = 0; i < 5; i += 1) {
    totalQty += toFloat(getItemValue(i, "count"));
    totalWeight += toFloat(getItemValue(i, "weight"));
    totalAmount += toFloat(getItemValue(i, "amount"));
  }
  return { totalQty, totalWeight, totalAmount };
}

async function exportInvoice() {
  await commitActiveInput();
  const items = gatherItems();
  if (!items.length) return alert("请至少填写一行产品");
  if (items.length > 5) return alert("产品行最多 5 行");

  const invoiceNo = toInt(qs("#invNo").value, 0);
  if (invoiceNo <= 0) return alert("序号必须为正整数");

  const customer = qs("#invCustomer").value.trim();
  let dateStr = formatDateInput(qs("#invDate").value.trim());
  qs("#invDate").value = dateStr;
  const location = qs("#invLocation").value.trim();
  const plate1 = qs("#invPlate1").value.trim();
  const plate2 = qs("#invPlate2").value.trim();
  const driver = qs("#invDriver").value.trim();
  const phone = qs("#invPhone").value.trim();
  const summary = qs("#invSummary").value.trim();

  const { totalQty, totalWeight, totalAmount } = getTotals();
  const category = resolveCategory(items);

  if (!qs("#invSummary").value.trim()) {
    qs("#invSummary").value = `共${totalQty.toFixed(0)}件，${totalWeight.toFixed(3)}吨，合计${totalAmount.toFixed(0)}元`;
  }

  const exportItems = items.map((it) => ({
    name: it.name,
    spec: it.spec,
    count: it.count,
    weight: it.weight,
    price: it.price,
    amount: it.amount,
  }));

  const filename = safeFilename(
    `${invoiceNo}-${category}-${dateStr.replace(/\//g, ".")}-${location}-${customer}-${totalQty.toFixed(0)}-${totalWeight.toFixed(3)}-${totalAmount.toFixed(0)}.docx`
  );

  const settings = {
    company_title: await getSetting("company_title", DEFAULT_COMPANY_TITLE),
    account_text: await getSetting("account_text", DEFAULT_ACCOUNT_TEXT),
  };

  await exportDocx({
    customer,
    date: dateStr,
    invoice_no: invoiceNo,
    location,
    plate1,
    plate2,
    driver,
    phone,
    summary,
    items: exportItems,
    total_qty: totalQty.toFixed(0),
    total_weight: totalWeight.toFixed(3),
    total_amount: totalAmount.toFixed(0),
  }, filename, settings);

  const dateIso = parseDateToIso(dateStr);
  const invoiceId = await addItem("invoices", {
    invoice_no: invoiceNo,
    customer,
    date: dateStr,
    date_iso: dateIso,
    category,
    location,
    total_qty: totalQty,
    total_weight: totalWeight,
    total_amount: totalAmount,
    filename,
    created_at: new Date().toISOString().slice(0, 19),
  });

  for (const it of exportItems) {
    await addItem("invoice_items", {
      invoice_id: invoiceId,
      product_name: it.name,
      spec_jin: toFloat(it.spec, null),
      qty: toFloat(it.count, null),
      weight: toFloat(it.weight, null),
      price: toFloat(it.price, null),
      amount: toFloat(it.amount, null),
    });
  }

  await addItem("invoice_audit", {
    invoice_id: invoiceId,
    action: "新增",
    detail: `导出单号${invoiceNo}`,
    created_at: new Date().toISOString().slice(0, 19),
  });

  await setSetting("last_invoice_no", String(invoiceNo));
  qs("#invNo").value = String(invoiceNo + 1);

  await autoSaveReferenceData(items, customer, location, plate1, plate2, driver, phone);
  await refreshReferenceData();
  await renderCustomers();
  await renderProducts();
  await renderVehicles();
  await statQuery(true);

  alert(`已导出：${filename}`);
}

async function autoSaveReferenceData(items, customer, location, plate1, plate2, driver, phone) {
  const customers = await getAll("customers");
  const products = await getAll("products");
  const vehicles = await getAll("vehicles");

  if (customer) {
    const nkey = normalizeText(customer);
    let existing = customers.find((c) => normalizeText(c.name) === nkey);
    if (!existing) {
      await addItem("customers", {
        name: customer,
        location,
        province: resolveProvince(location),
        note: "",
      });
    } else {
      const updated = {
        ...existing,
        location: existing.location || location || "",
        province: existing.province || resolveProvince(location),
      };
      await putItem("customers", updated);
    }
  }

  for (const it of items) {
    if (!it.name) continue;
    const nkey = normalizeText(it.name);
    const existing = products.find((p) => normalizeText(p.name) === nkey);
    if (!existing) {
      await addItem("products", {
        name: it.name,
        spec_jin: it.spec ? toFloat(it.spec, "") : "",
        category: "",
        note: "",
      });
    }
  }

  if (plate1 || driver || phone) {
    const existing = vehicles.find((v) => (v.plate1 || "") === plate1);
    if (!existing && plate1) {
      await addItem("vehicles", {
        plate1,
        plate2,
        driver,
        phone,
      });
    } else if (existing) {
      const updated = {
        ...existing,
        plate2: existing.plate2 || plate2 || "",
        driver: existing.driver || driver || "",
        phone: existing.phone || phone || "",
      };
      await putItem("vehicles", updated);
    }
  }
}

async function exportDocx(data, filename, settings) {
  const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    AlignmentType,
    Table,
    TableRow,
    TableCell,
    TableLayoutType,
    WidthType,
    BorderStyle,
    VerticalAlign,
    HeightRule,
    VerticalMergeType,
  } = docx;

  const vMergeRestart = VerticalMergeType ? VerticalMergeType.RESTART : undefined;
  const vMergeContinue = VerticalMergeType ? VerticalMergeType.CONTINUE : undefined;

  const fontName = "Microsoft YaHei";
  const setText = (text, size = 18, bold = false) => new TextRun({
    text: text == null ? "" : String(text),
    font: fontName,
    size,
    bold,
  });

  const cell = (text, opts = {}) => new TableCell({
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [setText(text, opts.size || 18, opts.bold || false)],
    })],
    width: opts.width,
    columnSpan: opts.colSpan,
    rowSpan: opts.rowSpan,
    verticalAlign: VerticalAlign.CENTER,
    borders: opts.borders,
    verticalMerge: opts.vMerge,
  });

  const dxaWidths = [1457, 1346, 1346, 1346, 1346, 1588, 1104, 1479];
  const totalDxa = dxaWidths.reduce((sum, v) => sum + v, 0);
  const rowHeight = { value: 284, rule: HeightRule.EXACT };

  const border = {
    top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  };

  const items = data.items || [];
  const maxRows = 5;

  const rows = [];

  rows.push(new TableRow({
    height: rowHeight,
    children: [
      cell(`客户：${data.customer || ""}`, { colSpan: 4, align: AlignmentType.LEFT, borders: border }),
      cell("销售日期", { borders: border }),
      cell(data.date || "", { borders: border }),
      cell("序号", { borders: border }),
      cell(String(data.invoice_no || ""), { borders: border }),
    ],
  }));

  rows.push(new TableRow({
    height: rowHeight,
    children: [
      cell("产品名称", { bold: true, borders: border }),
      cell("规格/斤", { bold: true, borders: border }),
      cell("件数", { bold: true, borders: border }),
      cell("重量/吨", { bold: true, borders: border }),
      cell("价格/吨", { bold: true, borders: border }),
      cell("金额", { bold: true, borders: border }),
      cell("到达地点", { bold: true, colSpan: 2, borders: border }),
    ],
  }));

  const priceNums = [];

  for (let i = 0; i < maxRows; i += 1) {
    const it = items[i] || {};
    const rawSpec = it.spec || "";
    const rawCount = it.count || "";
    const rawWeight = it.weight || "";
    const rawPrice = it.price || "";
    const rawAmount = it.amount || "";

    const specNum = toFloat(rawSpec, NaN);
    const countNum = toFloat(rawCount, NaN);
    const priceNum = toFloat(rawPrice, NaN);
    const amountNum = toFloat(rawAmount, NaN);

    let weightNum = toFloat(rawWeight, NaN);
    if (!Number.isFinite(weightNum) && Number.isFinite(specNum) && Number.isFinite(countNum)) {
      weightNum = (specNum * countNum) / 2000;
    }

    let amountCalc = Number.isFinite(amountNum) ? amountNum : NaN;
    if (!Number.isFinite(amountCalc) && Number.isFinite(weightNum) && Number.isFinite(priceNum)) {
      amountCalc = weightNum * priceNum;
    }

    let priceCalc = Number.isFinite(priceNum) ? priceNum : NaN;
    if (!Number.isFinite(priceCalc) && Number.isFinite(amountNum) && Number.isFinite(weightNum) && weightNum) {
      priceCalc = amountNum / weightNum;
    }

    if (Number.isFinite(priceNum)) priceNums.push(priceNum);
    else if (Number.isFinite(priceCalc)) priceNums.push(priceCalc);

    const dispSpec = rawSpec ? rawSpec : (Number.isFinite(specNum) ? formatDecimal(specNum) : "");
    const dispCount = rawCount ? rawCount : (Number.isFinite(countNum) ? formatDecimal(countNum) : "");
    const dispWeight = rawWeight ? formatDecimal(toFloat(rawWeight, 0)) : (Number.isFinite(weightNum) ? weightNum.toFixed(3) : "");
    const dispPrice = rawPrice ? rawPrice : (Number.isFinite(priceCalc) ? formatDecimal(priceCalc) : (Number.isFinite(priceNum) ? formatDecimal(priceNum) : ""));
    const dispAmount = rawAmount ? rawAmount : (Number.isFinite(amountCalc) ? Math.round(amountCalc).toString() : (Number.isFinite(amountNum) ? Math.round(amountNum).toString() : ""));

    rows.push(new TableRow({
      height: rowHeight,
      children: [
        cell(it.name || "", { borders: border }),
        cell(dispSpec, { borders: border }),
        cell(dispCount, { borders: border }),
        cell(dispWeight, { borders: border }),
        cell(dispPrice, { borders: border }),
        cell(dispAmount, { borders: border }),
        cell(i === 0 ? (data.location || "") : "", { colSpan: 2, borders: border }),
      ],
    }));
  }

  rows.push(new TableRow({
    height: rowHeight,
    children: [
      cell("", { borders: border }),
      cell("", { borders: border }),
      cell("", { borders: border }),
      cell("", { borders: border }),
      cell("", { borders: border }),
      cell("", { borders: border }),
      cell("车牌号1", { borders: border }),
      cell(data.plate1 || "", { borders: border }),
    ],
  }));

  let priceVal = "";
  if (priceNums.length && priceNums.every((p) => Math.abs(p - priceNums[0]) < 1e-9)) {
    priceVal = formatDecimal(priceNums[0]);
  }

  rows.push(new TableRow({
    height: rowHeight,
    children: [
      cell("总结", { bold: true, borders: border }),
      cell("", { bold: true, borders: border }),
      cell(data.total_qty || "", { bold: true, borders: border }),
      cell(data.total_weight || "", { bold: true, borders: border }),
      cell(priceVal, { bold: true, borders: border }),
      cell(data.total_amount || "", { bold: true, borders: border }),
      cell("车牌号2", { borders: border }),
      cell(data.plate2 || "", { borders: border }),
    ],
  }));

  const accountText = settings.account_text || DEFAULT_ACCOUNT_TEXT;
  const accountLines = accountText.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
  const account1 = accountLines[0] || accountText.trim();
  const account2 = accountLines[1] || "";

  rows.push(new TableRow({
    height: rowHeight,
    children: [
      cell("收款账号", { borders: border, vMerge: vMergeRestart }),
      cell(account1, { colSpan: 5, align: AlignmentType.LEFT, borders: border }),
      cell("车姓名", { borders: border }),
      cell(data.driver || "", { borders: border }),
    ],
  }));

  rows.push(new TableRow({
    height: rowHeight,
    children: [
      cell("", { borders: border, vMerge: vMergeContinue }),
      cell(account2, { colSpan: 5, align: AlignmentType.LEFT, borders: border }),
      cell("联系电话", { borders: border }),
      cell(data.phone || "", { borders: border }),
    ],
  }));

  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows,
    columnWidths: dxaWidths,
  });

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: 11339, height: 4535 },
            margin: { top: 227, bottom: 227, left: 340, right: 340 },
          },
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 0 },
            children: [setText(settings.company_title || DEFAULT_COMPANY_TITLE, 26, true)],
          }),
          table,
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  await downloadBlob(blob, filename);
}

async function downloadBlob(blob, filename) {
  const file = new File([blob], filename, { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try {
      await navigator.share({ files: [file], title: filename });
      return;
    } catch (err) {
      // fallback to normal download
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function statQuery(clearRegionFilter) {
  if (clearRegionFilter) state.statRegionFilterProv = null;
  state.statView = "query";
  const keyword = qs("#statKeyword").value.trim();
  const customer = qs("#statCustomer").value.trim();
  const location = qs("#statLocation").value.trim();

  let rows = await getAll("invoices");
  if (qs("#statMode").value !== "所有时间") {
    const [start, end] = getStatRange();
    rows = rows.filter((r) => r.date_iso && r.date_iso >= start && r.date_iso <= end);
  }

  rows = rows.filter((r) => {
    if (customer && !(r.customer || "").includes(customer)) return false;
    if (location && !(r.location || "").includes(location)) return false;
    if (keyword) {
      const blob = `${r.customer || ""}${r.location || ""}${r.invoice_no || ""}`;
      if (!blob.includes(keyword)) return false;
    }
    return true;
  });

  if (state.statRegionFilterProv) {
    rows = rows.filter((r) => resolveProvince(r.location || "") === state.statRegionFilterProv);
  }

  rows.sort((a, b) => String(b.date_iso || "").localeCompare(String(a.date_iso || "")));

  const columns = ["序号", "日期", "客户", "地点", "件数", "重量/吨", "金额/元"];
  renderStatTable(columns, rows.map((r) => ({
    id: r.id,
    values: [
      r.invoice_no || "",
      r.date || "",
      r.customer || "",
      r.location || "",
      r.total_qty != null ? toFloat(r.total_qty).toFixed(0) : "",
      r.total_weight != null ? `${toFloat(r.total_weight).toFixed(3)}吨` : "",
      r.total_amount != null ? `${toFloat(r.total_amount).toFixed(0)}元` : "",
    ],
  })));

  let totalQty = 0;
  let totalWeight = 0;
  let totalAmount = 0;
  rows.forEach((r) => {
    totalQty += toFloat(r.total_qty);
    totalWeight += toFloat(r.total_weight);
    totalAmount += toFloat(r.total_amount);
  });
  qs("#statSummary").textContent = `汇总：件数 ${totalQty.toFixed(0)}，重量 ${totalWeight.toFixed(2)}吨，金额 ${formatAmountCn(totalAmount)}`;
  await refreshAudit();
}

async function statCustomer() {
  state.statView = "customer";
  let items = await getAll("invoice_items");
  const invoices = await getAll("invoices");
  const invoiceMap = new Map(invoices.map((i) => [i.id, i]));

  if (qs("#statMode").value !== "所有时间") {
    const [start, end] = getStatRange();
    items = items.filter((it) => {
      const inv = invoiceMap.get(it.invoice_id);
      return inv && inv.date_iso && inv.date_iso >= start && inv.date_iso <= end;
    });
  }

  const totalAllAmount = items.reduce((s, r) => s + toFloat(r.amount), 0);
  const totalAllQty = items.reduce((s, r) => s + toFloat(r.qty), 0);
  const totalAllWeight = items.reduce((s, r) => s + toFloat(r.weight), 0);

  const agg = {};
  const invoiceIdsByCustomer = {};
  items.forEach((r) => {
    const inv = invoiceMap.get(r.invoice_id);
    const cust = (inv && inv.customer) || "未填";
    if (!agg[cust]) {
      agg[cust] = { qty: 0, weight: 0, amount: 0 };
      invoiceIdsByCustomer[cust] = new Set();
    }
    agg[cust].qty += toFloat(r.qty);
    agg[cust].weight += toFloat(r.weight);
    agg[cust].amount += toFloat(r.amount);
    invoiceIdsByCustomer[cust].add(r.invoice_id);
  });

  const columns = ["客户", "件数", "重量/吨", "金额/元", "次数", "件数占比", "重量占比", "金额占比"];
  const rows = Object.entries(agg).map(([cust, v]) => {
    const pctQty = totalAllQty ? (v.qty / totalAllQty) * 100 : 0;
    const pctWeight = totalAllWeight ? (v.weight / totalAllWeight) * 100 : 0;
    const pctAmount = totalAllAmount ? (v.amount / totalAllAmount) * 100 : 0;
    return {
      id: cust,
      values: [
        cust,
        v.qty.toFixed(0),
        `${v.weight.toFixed(3)}吨`,
        `${v.amount.toFixed(0)}元`,
        invoiceIdsByCustomer[cust].size,
        `${pctQty.toFixed(0)}%`,
        `${pctWeight.toFixed(0)}%`,
        `${pctAmount.toFixed(0)}%`,
      ],
    };
  });

  renderStatTable(columns, rows);
  qs("#statSummary").textContent = `汇总：件数 ${totalAllQty.toFixed(0)}，重量 ${totalAllWeight.toFixed(2)}吨，金额 ${formatAmountCn(totalAllAmount)}`;
}

async function statRegion() {
  state.statView = "region";
  let items = await getAll("invoice_items");
  const invoices = await getAll("invoices");
  const invoiceMap = new Map(invoices.map((i) => [i.id, i]));

  if (qs("#statMode").value !== "所有时间") {
    const [start, end] = getStatRange();
    items = items.filter((it) => {
      const inv = invoiceMap.get(it.invoice_id);
      return inv && inv.date_iso && inv.date_iso >= start && inv.date_iso <= end;
    });
  }

  const aggCity = {};
  const aggProv = {};
  let totalAmount = 0;

  items.forEach((r) => {
    const inv = invoiceMap.get(r.invoice_id);
    const city = (inv && inv.location) || "";
    const prov = resolveProvince(city);
    const amt = toFloat(r.amount);
    totalAmount += amt;
    aggCity[city] = (aggCity[city] || 0) + amt;
    aggProv[prov] = (aggProv[prov] || 0) + amt;
  });

  const mode = qs("#statRegionMode").value;
  const columns = ["地区", "金额/元", "占比"];
  const rows = [];
  const source = mode === "按省" ? aggProv : aggCity;
  Object.entries(source).forEach(([region, amt]) => {
    const pct = totalAmount ? (amt / totalAmount) * 100 : 0;
    rows.push({
      id: region,
      values: [region || "未知", `${amt.toFixed(0)}元`, `${pct.toFixed(2)}%`],
    });
  });

  renderStatTable(columns, rows);
  qs("#statSummary").textContent = `汇总：金额 ${formatAmountCn(totalAmount)}`;
}

function renderStatTable(columns, rows) {
  const thead = qs("#statTable thead");
  const tbody = qs("#statTable tbody");
  thead.innerHTML = `<tr>${columns.map((c) => `<th>${c}</th>`).join("")}</tr>`;
  tbody.innerHTML = rows
    .map((row) => `<tr data-id="${row.id}">${row.values.map((v) => `<td>${v}</td>`).join("")}</tr>`)
    .join("");
}

async function refreshAudit() {
  state.invoiceAudit = await getAll("invoice_audit");
  const rows = state.invoiceAudit
    .slice()
    .sort((a, b) => b.id - a.id)
    .slice(0, 50);
  const invoices = await getAll("invoices");
  const invoiceMap = new Map(invoices.map((i) => [i.id, i]));
  const tbody = qs("#auditTable tbody");
  tbody.innerHTML = rows
    .map((r) => {
      const inv = invoiceMap.get(r.invoice_id);
      return `<tr>
        <td>${r.created_at || ""}</td>
        <td>${r.action || ""}</td>
        <td>${inv ? inv.invoice_no || "" : ""}</td>
        <td>${inv ? inv.customer || "" : ""}</td>
        <td>${r.detail || ""}</td>
      </tr>`;
    })
    .join("");
}

function getStatRange() {
  const mode = qs("#statMode").value;
  const today = new Date();
  if (mode === "最近1个月") {
    return [toIso(addMonths(today, -1)), toIso(today)];
  }
  if (mode === "最近3个月") {
    return [toIso(addMonths(today, -3)), toIso(today)];
  }
  if (mode === "最近6个月") {
    return [toIso(addMonths(today, -6)), toIso(today)];
  }
  if (mode === "最近1年") {
    return [toIso(addMonths(today, -12)), toIso(today)];
  }
  if (mode === "按年") {
    const y = toInt(qs("#statYear").value, today.getFullYear());
    return [`${y}-01-01`, `${y}-12-31`];
  }
  if (mode === "按月") {
    const y = toInt(qs("#statYear").value, today.getFullYear());
    const m = toInt(qs("#statMonth").value, today.getMonth() + 1);
    const [s, e] = monthStartEnd(y, m);
    return [toIso(s), toIso(e)];
  }
  if (mode === "按季度") {
    const y = toInt(qs("#statYear").value, today.getFullYear());
    const q = toInt(qs("#statQuarter").value, Math.floor(today.getMonth() / 3) + 1);
    const [s, e] = quarterStartEnd(y, q);
    return [toIso(s), toIso(e)];
  }
  return [toIso(new Date(today.getFullYear(), today.getMonth(), 1)), toIso(today)];
}

function toIso(dateObj) {
  return dateObj.toISOString().slice(0, 10);
}

function bindStatsMode() {
  const updateMode = () => {
    const mode = qs("#statMode").value;
    const today = new Date();
    const yearInput = qs("#statYear");
    const monthSelect = qs("#statMonth");
    const quarterSelect = qs("#statQuarter");

    if (mode === "最近1个月" || mode === "最近3个月" || mode === "最近6个月" || mode === "最近1年") {
      yearInput.value = today.getFullYear();
      monthSelect.value = "";
      quarterSelect.value = "";
      yearInput.disabled = true;
      monthSelect.disabled = true;
      quarterSelect.disabled = true;
    } else if (mode === "按年") {
      yearInput.disabled = false;
      monthSelect.disabled = true;
      quarterSelect.disabled = true;
    } else if (mode === "按月") {
      yearInput.disabled = false;
      monthSelect.disabled = false;
      quarterSelect.disabled = true;
    } else if (mode === "按季度") {
      yearInput.disabled = false;
      monthSelect.disabled = true;
      quarterSelect.disabled = false;
    } else if (mode === "自己选择") {
      yearInput.disabled = false;
      monthSelect.disabled = false;
      quarterSelect.disabled = false;
    } else if (mode === "所有时间") {
      yearInput.value = "";
      monthSelect.value = "";
      quarterSelect.value = "";
      yearInput.disabled = true;
      monthSelect.disabled = true;
      quarterSelect.disabled = true;
    }
  };
  qs("#statMode").addEventListener("change", updateMode);
  updateMode();

  qs("#statMonth").addEventListener("change", () => {
    if (qs("#statMode").value === "自己选择" && qs("#statMonth").value) {
      qs("#statQuarter").value = "";
    }
  });

  qs("#statQuarter").addEventListener("change", () => {
    if (qs("#statMode").value === "自己选择" && qs("#statQuarter").value) {
      qs("#statMonth").value = "";
    }
  });
}

function getSelectedStatId() {
  const row = qs("#statTable tbody tr.selected");
  return row ? row.dataset.id : null;
}

function getSelectedStatIds() {
  return qsa("#statTable tbody tr.selected").map((r) => r.dataset.id);
}

function bindStatSelection() {
  let lastIndex = null;
  qs("#statTable").addEventListener("click", (e) => {
    const tr = e.target.closest("tr");
    if (!tr) return;
    const rows = qsa("#statTable tbody tr");
    const idx = rows.indexOf(tr);
    if (e.shiftKey && lastIndex !== null) {
      const [start, end] = idx > lastIndex ? [lastIndex, idx] : [idx, lastIndex];
      rows.forEach((r, i) => {
        if (i >= start && i <= end) {
          r.classList.add("selected");
        }
      });
    } else if (e.metaKey || e.ctrlKey) {
      tr.classList.toggle("selected");
      lastIndex = idx;
      return;
    } else {
      rows.forEach((r) => r.classList.remove("selected"));
      tr.classList.add("selected");
    }
    lastIndex = idx;
  });

  qs("#statTable").addEventListener("dblclick", () => {
    onStatDetail();
  });
}

async function onStatDetail() {
  const id = getSelectedStatId();
  if (!id) return alert("请选择一条记录");
  const inv = await getByKey("invoices", Number(id));
  if (!inv) return alert("记录不存在");
  const items = (await getAll("invoice_items")).filter((it) => it.invoice_id === inv.id);
  showDetailModal(inv, items);
}

async function onStatEdit() {
  const id = getSelectedStatId();
  if (!id) return alert("请选择一条记录");
  const inv = await getByKey("invoices", Number(id));
  if (!inv) return alert("记录不存在");
  showEditModal(inv);
}

async function onStatDelete() {
  const ids = getSelectedStatIds();
  if (!ids.length) return alert("请选择记录");
  if (!confirm(`确定要删除选中的 ${ids.length} 条记录吗？`)) return;
  const items = await getAll("invoice_items");
  for (const id of ids) {
    const invId = Number(id);
    const toDelete = items.filter((it) => it.invoice_id === invId);
    for (const it of toDelete) {
      await deleteItem("invoice_items", it.id);
    }
    await deleteItem("invoices", invId);
    await addItem("invoice_audit", {
      invoice_id: invId,
      action: "删除",
      detail: "删除记录",
      created_at: new Date().toISOString().slice(0, 19),
    });
  }
  await statQuery(true);
}

function showModal(title, bodyNode, actions = []) {
  qs("#modalTitle").textContent = title;
  const body = qs("#modalBody");
  body.innerHTML = "";
  body.appendChild(bodyNode);
  const actionsBox = qs("#modalActions");
  actionsBox.innerHTML = "";
  actions.forEach((btn) => actionsBox.appendChild(btn));
  qs("#modal").classList.remove("hidden");
}

function closeModal() {
  qs("#modal").classList.add("hidden");
}

function showDetailModal(inv, items) {
  const wrapper = document.createElement("div");
  const info = document.createElement("div");
  info.textContent = `序号：${inv.invoice_no}    日期：${inv.date}    客户：${inv.customer}    地点：${inv.location}\n件数：${toFloat(inv.total_qty).toFixed(0)}    重量：${inv.total_weight}    金额：${formatAmountCn(inv.total_amount)}`;
  info.style.whiteSpace = "pre-wrap";
  info.style.marginBottom = "10px";
  wrapper.appendChild(info);

  const table = document.createElement("table");
  table.className = "data-table";
  table.innerHTML = `
    <thead><tr><th>产品</th><th>规格/斤</th><th>件数</th><th>重量</th><th>价格/吨</th><th>金额/元</th></tr></thead>
    <tbody></tbody>
  `;
  const tbody = table.querySelector("tbody");
  items.forEach((it) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${it.product_name || ""}</td>
      <td>${it.spec_jin || ""}</td>
      <td>${it.qty || ""}</td>
      <td>${it.weight || ""}</td>
      <td>${it.price || ""}</td>
      <td>${it.amount != null ? `${toFloat(it.amount).toFixed(0)}元` : ""}</td>
    `;
    tbody.appendChild(tr);
  });
  wrapper.appendChild(table);

  const btn = document.createElement("button");
  btn.className = "primary";
  btn.textContent = "重新导出 Word";
  btn.addEventListener("click", async () => {
    const settings = {
      company_title: await getSetting("company_title", DEFAULT_COMPANY_TITLE),
      account_text: await getSetting("account_text", DEFAULT_ACCOUNT_TEXT),
    };
    const exportItems = items.map((it) => ({
      name: String(it.product_name || ""),
      spec: String(it.spec_jin || ""),
      count: String(it.qty || ""),
      weight: String(it.weight || ""),
      price: String(it.price || ""),
      amount: String(it.amount || ""),
    }));
    const safeDate = String(inv.date || "").replace(/\//g, ".");
    const category = inv.category || "鸡";
    const filename = safeFilename(
      `${inv.invoice_no}-${category}-${safeDate}-${inv.location}-${inv.customer}-${toFloat(inv.total_qty).toFixed(0)}-${inv.total_weight}-${toFloat(inv.total_amount).toFixed(0)}.docx`
    );
    await exportDocx({
      customer: inv.customer || "",
      date: inv.date || "",
      invoice_no: inv.invoice_no || "",
      location: inv.location || "",
      plate1: "",
      plate2: "",
      driver: "",
      phone: "",
      summary: "",
      items: exportItems,
      total_qty: toFloat(inv.total_qty).toFixed(0),
      total_weight: String(inv.total_weight || ""),
      total_amount: toFloat(inv.total_amount).toFixed(0),
    }, filename, settings);
  });

  showModal("订单详情", wrapper, [btn]);
}

function showEditModal(inv) {
  const wrapper = document.createElement("div");
  wrapper.innerHTML = `
    <div class="grid grid-3">
      <label class="field"><span>日期</span><input id="editDate" value="${inv.date || ""}" /></label>
      <label class="field"><span>客户</span><input id="editCustomer" value="${inv.customer || ""}" /></label>
      <label class="field"><span>地点</span><input id="editLocation" value="${inv.location || ""}" /></label>
      <label class="field"><span>件数</span><input id="editQty" value="${toFloat(inv.total_qty).toFixed(0)}" /></label>
      <label class="field"><span>重量</span><input id="editWeight" value="${inv.total_weight || ""}" /></label>
      <label class="field"><span>金额</span><input id="editAmount" value="${inv.total_amount || ""}" /></label>
    </div>
  `;

  const btnSave = document.createElement("button");
  btnSave.className = "primary";
  btnSave.textContent = "保存";
  btnSave.addEventListener("click", async () => {
    const updated = {
      ...inv,
      date: formatDateInput(qs("#editDate").value.trim()),
      date_iso: parseDateToIso(qs("#editDate").value.trim()),
      customer: qs("#editCustomer").value.trim(),
      location: qs("#editLocation").value.trim(),
      total_qty: toFloat(qs("#editQty").value),
      total_weight: toFloat(qs("#editWeight").value),
      total_amount: toFloat(qs("#editAmount").value),
    };
    await putItem("invoices", updated);
    await addItem("invoice_audit", {
      invoice_id: inv.id,
      action: "修改",
      detail: `修改单号${inv.invoice_no}`,
      created_at: new Date().toISOString().slice(0, 19),
    });
    closeModal();
    await statQuery(true);
  });

  showModal("修改记录", wrapper, [btnSave]);
}

async function exportJson() {
  const payload = {
    meta: { app: "kaidan-pwa", exported_at: new Date().toISOString() },
    settings: await getAll("settings"),
    customers: await getAll("customers"),
    products: await getAll("products"),
    vehicles: await getAll("vehicles"),
    invoices: await getAll("invoices"),
    invoice_items: await getAll("invoice_items"),
    invoice_audit: await getAll("invoice_audit"),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const filename = `kaidan-backup-${new Date().toISOString().slice(0, 10)}.json`;
  await downloadBlob(blob, filename);
}

async function importJson() {
  const file = qs("#importFile").files[0];
  if (!file) return alert("请选择 JSON 文件");
  const mode = qs("#importMode").value;
  const text = await file.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (err) {
    return alert("JSON 格式不正确");
  }

  if (mode === "replace") {
    await clearAllStores();
  }

  await importStoreData("settings", data.settings || [], mode, "key");
  await importStoreData("customers", data.customers || [], mode, "name");
  await importStoreData("products", data.products || [], mode, "name");
  await importStoreData("vehicles", data.vehicles || [], mode, "plate1");
  await importStoreData("invoices", data.invoices || [], mode, null);
  await importStoreData("invoice_items", data.invoice_items || [], mode, null);
  await importStoreData("invoice_audit", data.invoice_audit || [], mode, null);

  await refreshReferenceData();
  await renderCustomers();
  await renderProducts();
  await renderVehicles();
  await loadSettings();
  await statQuery(true);

  alert("导入完成");
}

async function importStoreData(storeName, rows, mode, uniqueKey) {
  if (!Array.isArray(rows)) return;
  const existing = await getAll(storeName);
  const byId = new Map(existing.map((r) => [r.id, r]));
  const byUnique = uniqueKey ? new Map(existing.map((r) => [r[uniqueKey], r])) : null;

  for (const row of rows) {
    if (!row) continue;
    if (mode === "merge") {
      if (row.id && byId.has(row.id)) continue;
      if (uniqueKey && row[uniqueKey] && byUnique.has(row[uniqueKey])) {
        const merged = { ...byUnique.get(row[uniqueKey]), ...row };
        await putItem(storeName, merged);
      } else {
        await putItem(storeName, row);
      }
    } else {
      await putItem(storeName, row);
    }
  }
}

async function clearAllStores() {
  await clearStore("settings");
  await clearStore("customers");
  await clearStore("products");
  await clearStore("vehicles");
  await clearStore("invoices");
  await clearStore("invoice_items");
  await clearStore("invoice_audit");
}

async function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    try {
      await navigator.serviceWorker.register("./sw.js");
    } catch (err) {
      console.warn("Service worker failed", err);
    }
  }
}

function bindActions() {
  qs("#btnExport").addEventListener("click", exportInvoice);
  qs("#btnClear").addEventListener("click", () => {
    qsa("#itemsTable input").forEach((input) => (input.value = ""));
    qs("#invCustomer").value = "";
    qs("#invLocation").value = "";
    qs("#invPlate1").value = "";
    qs("#invPlate2").value = "";
    qs("#invDriver").value = "";
    qs("#invPhone").value = "";
    qs("#invSummary").value = "";
    recalcTotals();
  });

  qs("#cusSave").addEventListener("click", saveCustomer);
  qs("#cusDelete").addEventListener("click", deleteCustomer);
  qs("#cusClear").addEventListener("click", () => clearForm("#tab-customers"));

  qs("#proSave").addEventListener("click", saveProduct);
  qs("#proDelete").addEventListener("click", deleteProduct);
  qs("#proClear").addEventListener("click", () => clearForm("#tab-products"));

  qs("#vehSave").addEventListener("click", saveVehicle);
  qs("#vehDelete").addEventListener("click", deleteVehicle);
  qs("#vehClear").addEventListener("click", () => clearForm("#tab-vehicles"));

  qs("#btnStatQuery").addEventListener("click", () => statQuery(true));
  qs("#btnStatCustomer").addEventListener("click", statCustomer);
  qs("#btnStatRegion").addEventListener("click", statRegion);
  qs("#btnStatDetail").addEventListener("click", onStatDetail);
  qs("#btnStatEdit").addEventListener("click", onStatEdit);
  qs("#btnStatDelete").addEventListener("click", onStatDelete);

  qs("#btnSaveSettings").addEventListener("click", saveSettings);
  qs("#btnSaveAiSettings").addEventListener("click", saveAiSettings);
  qs("#btnTestAi").addEventListener("click", testAiApi);
  qs("#btnExportJson").addEventListener("click", exportJson);
  qs("#btnImport").addEventListener("click", importJson);
  qs("#btnClearAll").addEventListener("click", async () => {
    if (!confirm("确定清空全部数据吗？此操作不可恢复。")) return;
    await clearAllStores();
    await ensureDefaults();
    await refreshReferenceData();
    await renderCustomers();
    await renderProducts();
    await renderVehicles();
    await loadSettings();
    await statQuery(true);
  });

  qs("#modalClose").addEventListener("click", closeModal);
  qs("#modal").addEventListener("click", (e) => {
    if (e.target.id === "modal") closeModal();
  });

  bindTableSelection("#cusTable", (id) => {
    const item = state.customers.find((c) => c.id === Number(id));
    if (!item) return;
    qs("#cusName").value = item.name || "";
    qs("#cusLocation").value = item.location || "";
    qs("#cusProvince").value = item.province || "";
    qs("#cusNote").value = item.note || "";
  });

  bindTableSelection("#proTable", (id) => {
    const item = state.products.find((p) => p.id === Number(id));
    if (!item) return;
    qs("#proName").value = item.name || "";
    qs("#proSpec").value = item.spec_jin || "";
    qs("#proCategory").value = item.category || "鸡";
    qs("#proNote").value = item.note || "";
  });

  bindTableSelection("#vehTable", (id) => {
    const item = state.vehicles.find((v) => v.id === Number(id));
    if (!item) return;
    qs("#vehPlate1").value = item.plate1 || "";
    qs("#vehPlate2").value = item.plate2 || "";
    qs("#vehDriver").value = item.driver || "";
    qs("#vehPhone").value = item.phone || "";
  });

  bindMultiSelection("#cusTable");
  bindMultiSelection("#proTable");
  bindMultiSelection("#vehTable");

  qs("#btnAiImport").addEventListener("click", () => {
    qs("#aiImageInput").click();
  });
  qs("#aiImageInput").addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    try {
      await aiImportAndExport(file);
    } catch (err) {
      console.error(err);
      alert(err.message || "AI 识图失败");
      setAiStatus("识别失败");
    } finally {
      e.target.value = "";
    }
  });

  qs("#cusLocation").addEventListener("input", () => {
    const loc = qs("#cusLocation").value.trim();
    if (!qs("#cusProvince").value.trim()) {
      qs("#cusProvince").value = resolveProvince(loc);
    }
  });
}

async function init() {
  initTabs();
  bindInvoiceInputs();
  bindActions();
  bindStatsMode();
  bindStatSelection();

  state.db = await openDb();
  await ensureDefaults();
  await refreshReferenceData();
  await renderCustomers();
  await renderProducts();
  await renderVehicles();
  await loadSettings();
  await statQuery(true);
  recalcTotals();

  await registerServiceWorker();
  setStatus("已就绪");
}

init().catch((err) => {
  console.error(err);
  setStatus("加载失败");
});


