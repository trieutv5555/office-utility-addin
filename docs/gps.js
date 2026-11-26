const GPS = (() => {
'use strict';
function calculate(data, config){
  config = validateConfig(config);
  const FIRST_DATA_ROW = 5;

  const MIN_FACTORY_STOP_MINUTES = config.factoryMinutes;
  const MIN_REST_MINUTES = config.restMinutes;
  const MIN_TRIP_KM = config.minKm;
  const MAX_MERGE_GAP_MINUTES = 10;
  const MAX_REST_MOVE_KM = 1;
  const MAX_SAME_PLACE_DISTANCE_KM = 0.3;

  const times = data.map(row => excelDate(row[0]));

  // 1. Nhận diện nhà máy bằng GPS, không dựa vào tên địa chỉ.
  const factories = config.factories;
  const factoryCodes = data.map(row => {
    if (!validGps(row)) return "";
    for (const f of factories) {
      if (haversineKm(f.lat, f.lon, toNumber(row[2]),
        toNumber(row[1])) <= f.radius) return f.code;
    }
    return "";
  });

  // Nối các mẫu GPS lỗi ngắn, chỉ khi hai đầu cùng nhà máy,
  // xe đứng yên trong khoảng lỗi và đồng hồ km không tăng.
  for (let i = 0; i < data.length;) {
    if (validGps(data[i])) { i++; continue; }
    const start = i;
    while (i < data.length && !validGps(data[i])) i++;
    const before = start - 1;
    if (before < 0 || i >= data.length) continue;
    const code = factoryCodes[before];
    if (!code || code !== factoryCodes[i]) continue;
    const gap = (times[i] - times[before]) * 1440;
    const km = toNumber(data[i][4]) - toNumber(data[before][4]);
    if (gap < 0 || gap >= MAX_MERGE_GAP_MINUTES ||
        !Number.isFinite(km) || Math.abs(km) > 0.001) continue;
    let stationary = true;
    for (let j = start; j < i; j++) {
      if (toNumber(data[j][3]) !== 0) stationary = false;
    }
    if (stationary) {
      for (let j = start; j < i; j++) factoryCodes[j] = code;
    }
  }

  // 2. Một lần lưu tại nhà máy bao gồm cả di chuyển nội bộ.
  // Mốc đến: đầu lần dừng dài đầu tiên đạt ngưỡng cấu hình, nếu có.
  // Mốc đi: mẫu vận tốc 0 cuối cùng trước khi rời vùng GPS.
  const factoryStops = [];
  for (let i = 0; i < data.length;) {
    const code = factoryCodes[i];
    if (!code) { i++; continue; }
    const visitStart = i;
    while (i < data.length && factoryCodes[i] === code) i++;
    const visitEnd = i - 1;
    const runs = [];
    let stoppedAt = -1;
    for (let j = visitStart; j <= visitEnd + 1; j++) {
      const stopped = j <= visitEnd && toNumber(data[j][3]) === 0;
      if (stopped) {
        if (stoppedAt < 0) stoppedAt = j;
      } else if (stoppedAt >= 0) {
        runs.push({ start: stoppedAt, end: j - 1 });
        stoppedAt = -1;
      }
    }
    if (!runs.length) continue;
    const duration = (r) =>
      (times[r.end] - times[r.start]) * 1440;
    const totalStopped = runs.reduce((sum, r) => sum + duration(r), 0);
    if (totalStopped + 0.000001 < MIN_FACTORY_STOP_MINUTES) continue;
    const firstLong = runs.find(r =>
      duration(r) + 0.000001 >= MIN_FACTORY_STOP_MINUTES);
    factoryStops.push({
      code,
      start: firstLong ? firstLong.start : runs[0].start,
      end: runs[runs.length - 1].end
    });
  }

  // 3. Tạo cả các đoạn ở hai đầu dữ liệu.
  const segments = [];

  if (factoryStops.length) {
    const firstFactory = factoryStops[0];

    if (firstFactory.start > 0) {
      segments.push({
        start: 0,
        end: firstFactory.start,
        from: "",
        to: firstFactory.code,
        partial: "đầu"
      });
    }

    for (let i = 0; i < factoryStops.length - 1; i++) {
      segments.push({
        start: factoryStops[i].end,
        end: factoryStops[i + 1].start,
        from: factoryStops[i].code,
        to: factoryStops[i + 1].code,
        partial: ""
      });
    }

    const lastFactory = factoryStops[factoryStops.length - 1];

    if (lastFactory.end < data.length - 1) {
      segments.push({
        start: lastFactory.end,
        end: data.length - 1,
        from: lastFactory.code,
        to: "",
        partial: "cuối"
      });
    }
  }

  const trips = [];
  const rests = [];

  for (const segment of segments) {
    const departure = segment.start;
    const arrival = segment.end;

    const tripKm =
      toNumber(data[arrival][4]) -
      toNumber(data[departure][4]);

    // Ngưỡng quãng đường cấu hình áp dụng cho cả chuyến đủ mốc
    // lẫn chuyến thiếu đầu hoặc chưa về.
    if (!Number.isFinite(tripKm) || tripKm + 1e-9 < MIN_TRIP_KM) {
      continue;
    }

    const tripNumber = trips.length + 1;
    const outputRow = tripNumber + 4;
    const departureRow = departure + FIRST_DATA_ROW;
    const arrivalRow = arrival + FIRST_DATA_ROW;

    // Chuyến thiếu đầu: dùng GPS tại nhà máy về làm
    // gốc đo điểm xa nhất, vì không biết điểm xuất phát thật.
    const origin = segment.partial === "đầu"
      ? data[arrival]
      : data[departure];

    let farthest = -1;
    let maxDistance = -1;

    if (validGps(origin)) {
      for (let i = departure; i <= arrival; i++) {
        if (
          !validGps(data[i]) ||
          !String(data[i][5]).trim()
        ) {
          continue;
        }

        const distance = haversineKm(
          toNumber(origin[2]),
          toNumber(origin[1]),
          toNumber(data[i][2]),
          toNumber(data[i][1])
        );

        if (distance > maxDistance) {
          maxDistance = distance;
          farthest = i;
        }
      }
    }

    const farthestRow = farthest >= 0
      ? farthest + FIRST_DATA_ROW
      : 0;

    const passedFactories = [];

    for (let i = departure + 1; i < arrival; i++) {
      const code = factoryCodes[i];

      if (
        code &&
        code !== segment.from &&
        code !== segment.to &&
        !passedFactories.includes(code)
      ) {
        passedFactories.push(code);
      }
    }

    let note = "";

    if (segment.partial === "đầu") {
      note =
        "Chuyến đã bắt đầu trước dữ liệu; điểm đi là " +
        "vị trí GPS đầu. Km và giờ chỉ tính phần quan sát. " +
        "Điểm xa nhất tính từ nơi về.";
    } else if (segment.partial === "cuối") {
      note =
        "Chuyến chưa về khi hết dữ liệu; điểm về " +
        "là vị trí GPS cuối. Km và giờ chỉ tính phần quan sát.";
    } else if (segment.from !== segment.to) {
      note =
        `Chuyển từ ${segment.from} sang ${segment.to}; ` +
        "điểm xa nhất có thể gần điểm về.";
    }

    if (farthest < 0) {
      note +=
        (note ? " " : "") +
        "Cần kiểm tra: không xác định được GPS xa nhất.";
    }

    if (passedFactories.length) {
      note +=
        (note ? " " : "") +
        `Đi ngang ${passedFactories.join(", ")}: ` +
        `không dừng đủ ${MIN_FACTORY_STOP_MINUTES} phút.`;
    }

    const tripLabel = segment.partial === "cuối"
      ? `${tripNumber} (chưa về)`
      : tripNumber;

    const pointOfDeparture = segment.partial === "đầu"
      ? `=Data!G${departureRow}`
      : segment.from;

    const pointOfReturn = segment.partial === "cuối"
      ? `=Data!G${arrivalRow}`
      : segment.to;

    trips.push([
      tripLabel,                                  // A Chuyến
      pointOfDeparture,                           // B Điểm đi
      farthestRow ? `=Data!G${farthestRow}` : "", // C Điểm đến
      pointOfReturn,                              // D Điểm về
      times[departure],                           // E Bắt đầu
      times[arrival],                             // F Kết thúc
      `=L${outputRow}-K${outputRow}`,              // G Km chuyến
      `=(F${outputRow}-E${outputRow})*24`,         // H Số giờ
      departureRow,                               // I Dòng đầu
      arrivalRow,                                 // J Dòng cuối
      `=Data!F${departureRow}`,                   // K Km đầu
      `=Data!F${arrivalRow}`,                     // L Km cuối
      farthest >= 0 ? times[farthest] : "",        // M Lúc xa nhất
      farthestRow || "",                          // N Dòng xa nhất
      note                                        // O Ghi chú
    ]);

    // 4. Tìm và gộp các lần dừng nghỉ trong chuyến.
    const rawRestRuns = [];
    let restStart = -1;

    const scanStart = segment.partial === "đầu"
      ? departure
      : departure + 1;

    for (let i = scanStart; i <= arrival + 1; i++) {
      const stopped =
        i <= arrival &&
        (segment.partial === "cuối" || i < arrival) &&
        String(data[i][3]).trim() !== "" &&
        toNumber(data[i][3]) === 0;

      if (stopped) {
        if (restStart < 0) restStart = i;
      } else if (restStart >= 0) {
        rawRestRuns.push({
          start: restStart,
          end: i - 1
        });
        restStart = -1;
      }
    }

    const restGroups = [];

    for (const run of rawRestRuns) {
      const group = restGroups[restGroups.length - 1];
      const previous = group?.[group.length - 1];
      let shouldMerge = false;

      if (previous) {
        const firstRun = group[0];

        const gapMinutes =
          (times[run.start] - times[previous.end]) * 1440;

        const gapKm =
          toNumber(data[run.start][4]) -
          toNumber(data[previous.end][4]);

        const firstAddress =
          String(data[firstRun.start][5]).trim();

        const nextAddress =
          String(data[run.start][5]).trim();

        const sameAddress =
          firstAddress !== "" &&
          firstAddress === nextAddress;

        const nearbyGps =
          validGps(data[firstRun.start]) &&
          validGps(data[run.start]) &&
          haversineKm(
            toNumber(data[firstRun.start][2]),
            toNumber(data[firstRun.start][1]),
            toNumber(data[run.start][2]),
            toNumber(data[run.start][1])
          ) <= MAX_SAME_PLACE_DISTANCE_KM;

        shouldMerge =
          gapMinutes >= 0 &&
          gapMinutes < MAX_MERGE_GAP_MINUTES &&
          gapKm >= 0 &&
          gapKm < MAX_REST_MOVE_KM &&
          (sameAddress || nearbyGps);
      }

      if (shouldMerge) {
        group.push(run);
      } else {
        restGroups.push([run]);
      }
    }

    let orderInTrip = 0;

    for (const group of restGroups) {
      const first = group[0];
      const last = group[group.length - 1];

      const minutes =
        (times[last.end] - times[first.start]) * 1440;

      if (minutes <= MIN_REST_MINUTES + 1e-6) continue;

      rests.push({
        trip: tripNumber,
        order: ++orderInTrip,
        address:
          String(data[first.start][5] || "").trim() ||
          "Không có địa chỉ",
        start: first.start,
        end: last.end,
        minutes
      });
    }
  }


return {trips,rests,times};
}
function toNumber(value) {
  if (typeof value === "number") return value;
  const text = String(value).trim();
  return text === "" ? NaN : Number(text.replace(",", "."));
}

// Chuyển dd/mm/yyyy hh:mm:ss sang số ngày Excel.
function excelDate(value) {
  if (typeof value === "number") return value;

  const match = String(value).trim().match(
    /^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/
  );

  if (!match) {
    throw new Error(`Ngày giờ không đúng định dạng: ${value}`);
  }

  return Date.UTC(
    Number(match[3]),
    Number(match[2]) - 1,
    Number(match[1]),
    Number(match[4]),
    Number(match[5]),
    Number(match[6])
  ) / 86400000 + 25569;
}

function validGps(row) {
  const lon = toNumber(row[1]);
  const lat = toNumber(row[2]);

  return (
    Number.isFinite(lon) &&
    Number.isFinite(lat) &&
    lon >= -180 && lon <= 180 &&
    lat >= -90 && lat <= 90 &&
    !(lon === 0 && lat === 0)
  );
}

function haversineKm(
  lat1,
  lon1,
  lat2,
  lon2
) {
  const radians = Math.PI / 180;
  const dLat = (lat2 - lat1) * radians;
  const dLon = (lon2 - lon1) * radians;

  const h =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1 * radians) *
    Math.cos(lat2 * radians) *
    Math.sin(dLon / 2) ** 2;

  return 12742 *
    Math.asin(Math.sqrt(Math.min(1, Math.max(0, h))));
}
// Local device configuration only. No requests are made with these values.
const LEGACY_KEY='office-utility.gps.settings.v1';
const KEY='office-utility.gps.settings.v2';
const MAX_FACTORIES=50;
const root=()=>document.getElementById('gps-factories');
const readNum=v=>{const s=String(v).trim().replace(',','.');return s===''?NaN:Number(s);};
function positive(value,max,label){
 if(!Number.isFinite(value)||value<=0||value>max)throw Error(label+' phải lớn hơn 0 và không vượt '+max+'.');
 return value;
}
function validateConfig(c){
 if(!c||!Array.isArray(c.factories)||c.factories.length<1||c.factories.length>MAX_FACTORIES)throw Error('Cần cấu hình từ 1 đến '+MAX_FACTORIES+' nhà máy.');
 const minKm=positive(c.minKm,100000,'Quãng đường tối thiểu (km)');
 const factoryMinutes=positive(c.factoryMinutes,10080,'Thời gian dừng tối thiểu tại nhà máy (phút)');
 const restMinutes=positive(c.restMinutes,10080,'Ngưỡng thống kê dừng xe giữa chuyến (phút)');
 const names=new Set();
 const factories=c.factories.map((f,i)=>{
  const code=typeof f?.code==='string'?f.code.trim().normalize('NFC'):'';
  if(!code||code.length>100||/^[=+@-]|[\x00-\x1f\x7f]/.test(code))throw Error('Tên nhà máy '+(i+1)+' không hợp lệ.');
  const key=code.toLocaleLowerCase('vi');if(names.has(key))throw Error('Tên nhà máy bị trùng: '+code);names.add(key);
  if(!Number.isFinite(f.lon)||!Number.isFinite(f.lat)||f.lon< -180||f.lon>180||f.lat< -90||f.lat>90||(f.lon===0&&f.lat===0))throw Error('Tọa độ của '+code+' không hợp lệ.');
  return {code,lon:f.lon,lat:f.lat,radius:positive(f.radius,100,'Bán kính của '+code+' (km)')};
 });
 for(let i=0;i<factories.length;i++)for(let j=i+1;j<factories.length;j++){
  const a=factories[i],b=factories[j];
  if(haversineKm(a.lat,a.lon,b.lat,b.lon)<=a.radius+b.radius)throw Error('Vùng GPS của '+a.code+' và '+b.code+' chồng lấn hoặc tiếp xúc. Kiểm tra tọa độ hoặc giảm bán kính.');
 }
 return {factories,minKm,factoryMinutes,restMinutes};
}
function migrateLegacy(old){
 if(!old||!Array.isArray(old.factories))throw Error('Cấu hình cũ không hợp lệ.');
 return validateConfig({factories:old.factories.map((f,i)=>({code:f.code||'NM'+(i+1),lon:f.lon,lat:f.lat,radius:old.radius})),minKm:old.minKm,factoryMinutes:30,restMinutes:10});
}
function rowField(row,key){return row.querySelector('[data-gps-field="'+key+'"]');}
function refreshRows(){
 const rows=Array.from(root().children);
 rows.forEach((row,i)=>{
  row.querySelector('.gps-factory-title').textContent='Nhà máy '+(i+1);
  row.querySelector('.gps-remove').disabled=rows.length===1;
 });
 const addBtn=document.querySelector('[data-action="xl-gps-add"]');
 if(addBtn)addBtn.setAttribute('aria-label','Thêm nhà máy (tối đa '+MAX_FACTORIES+')');
}
function appendFactory(f){
 const row=document.createElement('div');row.className='gps-factory';
 const head=document.createElement('div');head.className='gps-factory-head';
 const title=document.createElement('strong');title.className='gps-factory-title';
 const remove=document.createElement('button');remove.type='button';remove.className='gps-remove';remove.textContent='Xóa';
 remove.addEventListener('click',()=>{if(typeof busy!=='undefined'&&busy)return;if(root().children.length<=1)return;row.remove();refreshRows();});
 head.append(title,remove);row.append(head);
 const fields=[['code','Tên nhà máy',f.code],['lon','Kinh độ',f.lon],['lat','Vĩ độ',f.lat],['radius','Bán kính nhận diện (km)',f.radius]];
 const grid=document.createElement('div');grid.className='grid';
 for(const [key,label,value] of fields){
  const wrap=document.createElement('label');wrap.textContent=label;wrap.className=(key==='code'||key==='radius')?'wide':'';
  const input=document.createElement('input');input.type='text';input.dataset.gpsField=key;input.autocomplete='off';input.spellcheck=false;
  if(key!=='code')input.inputMode='decimal';input.value=value==null?'':String(value);
  wrap.append(input);
  if(key==='radius'){const hint=document.createElement('span');hint.className='gps-field-hint';hint.textContent='Phạm vi tính từ tọa độ trên; xe trong vùng này có thể được xác nhận đang ở nhà máy.';wrap.append(hint);}
  grid.append(wrap);
 }
 row.append(grid);root().append(row);refreshRows();
}
function add(){
 if(root().children.length>=MAX_FACTORIES)throw Error('Tối đa '+MAX_FACTORIES+' nhà máy.');
 const names=new Set(Array.from(root().children,r=>rowField(r,'code').value.trim().toLocaleLowerCase('vi')));
 let i=1;while(names.has(('NM'+i).toLowerCase()))i++;
 appendFactory({code:'NM'+i,radius:1});return 'Đã thêm nhà máy. Nhập tọa độ rồi lưu nếu cần.';
}
function readConfig(){
 const factories=Array.from(root().children,row=>({code:rowField(row,'code').value,lon:readNum(rowField(row,'lon').value),lat:readNum(rowField(row,'lat').value),radius:readNum(rowField(row,'radius').value)}));
 return validateConfig({factories,minKm:readNum(document.getElementById('gps-min-km').value),factoryMinutes:readNum(document.getElementById('gps-factory-minutes').value),restMinutes:readNum(document.getElementById('gps-rest-minutes').value)});
}
function displayConfig(c){
 root().replaceChildren();
 const factories=c?c.factories:[{code:'NM1',radius:1}];
 factories.forEach(appendFactory);
 document.getElementById('gps-min-km').value=String(c?c.minKm:5);
 document.getElementById('gps-factory-minutes').value=String(c?c.factoryMinutes:30);
 document.getElementById('gps-rest-minutes').value=String(c?c.restMinutes:10);
}
let savedSnapshot=null;
function showSettings(open,c=null){
 const panel=document.getElementById('gps-settings-panel');
 const summary=document.getElementById('gps-config-summary');
 const toggle=document.getElementById('gps-settings-toggle');
 panel.hidden=!open;
 summary.hidden=open;
 toggle.textContent=open?'Thu gọn cài đặt':'Chỉnh sửa cài đặt GPS';
 toggle.setAttribute('aria-expanded',String(open));
 if(!open&&c){
  const saved=JSON.stringify(c)===savedSnapshot;
  document.getElementById('gps-summary-state').textContent=saved?'Đã lưu cài đặt':'Cài đặt chưa lưu – dùng cho phiên này';
  document.getElementById('gps-summary-sites').textContent=c.factories.length+' nhà máy đã cấu hình';
  document.getElementById('gps-summary-rules').textContent='Chuyến ≥ '+c.minKm+' km · Về nhà máy ≥ '+c.factoryMinutes+' phút · Dừng giữa chuyến > '+c.restMinutes+' phút';
 }
}
function toggleSettings(){
 const panel=document.getElementById('gps-settings-panel');
 if(panel.hidden){showSettings(true);return 'Đã mở cài đặt GPS.';}
 const c=readConfig();showSettings(false,c);
 return JSON.stringify(c)===savedSnapshot?'Đã thu gọn cài đặt GPS.':'Đã thu gọn. Thay đổi chưa lưu, chỉ áp dụng trong phiên hiện tại.';
}
function init(){
 let c=null;
 try{
  const v=localStorage.getItem(KEY);
  if(v)c=validateConfig(JSON.parse(v));
  else{const old=localStorage.getItem(LEGACY_KEY);if(old){c=migrateLegacy(JSON.parse(old));try{localStorage.setItem(KEY,JSON.stringify(c));localStorage.removeItem(LEGACY_KEY);}catch(_){/* retain old settings if migration cannot persist */}}}
 }catch(_){c=null;}
 displayConfig(c);
 savedSnapshot=c?JSON.stringify(c):null;
 showSettings(!c,c);
}
function save(){
 const c=readConfig();
 try{localStorage.setItem(KEY,JSON.stringify(c));localStorage.removeItem(LEGACY_KEY);}catch(_){throw Error('Không lưu được hoàn toàn cấu hình cục bộ. Bạn vẫn có thể chạy trong phiên hiện tại.');}
 savedSnapshot=JSON.stringify(c);showSettings(false,c);
 return 'Đã lưu '+c.factories.length+' nhà máy và thu gọn cài đặt GPS.';
}
function resetDefaults(){
 document.getElementById('gps-min-km').value='5';
 document.getElementById('gps-factory-minutes').value='30';
 document.getElementById('gps-rest-minutes').value='10';
 return 'Đã khôi phục 3 điều kiện mặc định (5 km, 30 phút, 10 phút). Tên, tọa độ và bán kính từng nhà máy giữ nguyên. Bấm Lưu cài đặt nếu muốn lưu.';
}
async function clear(){
 if(!await officeConfirm('Xóa toàn bộ cấu hình GPS đã lưu trên thiết bị này, gồm cả tên và tọa độ nhà máy?'))return 'Đã huỷ.';
 try{localStorage.removeItem(KEY);localStorage.removeItem(LEGACY_KEY);}catch(_){throw Error('Không xóa được cấu hình trong bộ nhớ cục bộ.');}
 savedSnapshot=null;displayConfig(null);showSettings(true);return 'Đã xóa cấu hình GPS trên thiết bị này.';
}
// Office.js columnWidth is measured in points, not pixels.
function formatReport(out,{total,dt,dh,df,de,tripCount,restCount}){
 const all=out.getRange('A1:O'+de).format;
 all.font.name='Calibri';all.font.size=10;
 all.verticalAlignment='Center';all.wrapText=false;all.rowHeight=15;
 // The two report tables share columns; D/E/F must also accommodate date/time.
 const widths=[43,62,150,104,104,104,70,70,65,65,65,65,104,65,180];
 widths.forEach((w,i)=>{out.getRangeByIndexes(0,i,1,1).format.columnWidth=w;});
 for(const a of ['A4:O4','A'+dh+':H'+dh]){
  const f=out.getRange(a).format;f.wrapText=true;f.horizontalAlignment='Center';f.verticalAlignment='Center';
 }
 if(tripCount){
  const last=total-1;
  // Factory names, addresses and notes may be long; do not enlarge the columns.
  out.getRange('B5:D'+last).format.wrapText=true;
  out.getRange('O5:O'+last).format.wrapText=true;
  out.getRange('A5:A'+last).format.wrapText=true;
  out.getRange('A5:A'+last).format.horizontalAlignment='Center';
 }
 if(restCount){
  out.getRange('C'+df+':C'+(de-1)).format.wrapText=true;
  out.getRange('A'+df+':B'+(de-1)).format.horizontalAlignment='Center';
 }
 // Titles and full-width notes remain on a single line across empty cells.
 for(const a of ['A1:O1','A'+dt+':H'+dt])out.getRange(a).format.rowHeight=19;
 // Fit height only, after the widths and Wrap Text have been set.
 out.getRange('A4:O'+Math.max(4,total-1)).format.autofitRows();
 out.getRange('A'+dh+':H'+Math.max(dh,de-1)).format.autofitRows();
}
async function run(){
 const config=readConfig();
 if(!Office.context.requirements.isSetSupported('ExcelApi','1.4'))throw Error('Cần ExcelApi 1.4.');
 return Excel.run(async ctx=>{
 const sheets=ctx.workbook.worksheets,source=sheets.getItemOrNullObject('Data');await ctx.sync();
 if(source.isNullObject)throw Error('Không tìm thấy sheet Data.');
 const used=source.getRange('B:G').getUsedRangeOrNullObject(true);used.load('rowIndex,rowCount');await ctx.sync();
 if(used.isNullObject||used.rowIndex+used.rowCount<5)throw Error('Không có dữ liệu GPS từ dòng 5.');
 const last=used.rowIndex+used.rowCount;if(last>100004)throw Error('Tối đa 100.000 bản ghi GPS.');
 const data=[];for(let n=5;n<=last;n+=5000){const r=source.getRange('B'+n+':G'+Math.min(last,n+4999));r.load('values');await ctx.sync();data.push(...r.values);}
 const plate=source.getRange('A3');plate.load('values');await ctx.sync();
 const {trips,rests,times}=calculate(data, config);
 const name='Chuyến đi 1';let out=sheets.getItemOrNullObject(name);await ctx.sync();
 if(!out.isNullObject){const title=out.getRange('A1');title.load('values');await ctx.sync();if(!String(title.values[0][0]).startsWith('TÁCH CHUYẾN ĐI'))throw Error('Sheet Chuyến đi 1 đã tồn tại nhưng không phải báo cáo GPS. Hãy đổi tên sheet đó.');if(!await officeConfirm('Cập nhật báo cáo GPS trong sheet Chuyến đi 1?'))return 'Đã huỷ.';const old=out.getUsedRangeOrNullObject(true);old.load('rowIndex,rowCount');await ctx.sync();if(!old.isNullObject)out.getRange('A1:O'+Math.max(10,old.rowIndex+old.rowCount)).clear('All');}else out=sheets.add(name);
 const total=trips.length+5,dt=total+2,dh=total+5,df=dh+1,de=df+Math.max(rests.length,1);
 const matrix=Array.from({length:de},()=>Array(15).fill(''));
 const put=(r,values)=>values.forEach((v,c)=>matrix[r-1][c]=v);
 put(1,['TÁCH CHUYẾN ĐI – '+String(plate.values[0][0]).split(':').pop().trim()]);
 put(2,['Nhà máy: '+config.factories.map(f=>f.code+' ('+f.radius+' km)').join(' | ')]);put(3,['GPS – Chuyến ≥'+config.minKm+' km; xác nhận mốc nhà máy khi tổng thời gian dừng ≥'+config.factoryMinutes+' phút.']);
 put(4,['Chuyến','Điểm đi','Điểm xa nhất (điểm đến)','Điểm về','Bắt đầu','Kết thúc','Quãng đường (km)','Thời gian (giờ)','Dòng bắt đầu','Dòng kết thúc','Km đầu','Km cuối','Đến điểm xa nhất lúc','Dòng điểm xa nhất','Ghi chú']);
 trips.forEach((t,i)=>put(i+5,t.map((v,c)=>[1,2,3,6,7,10,11].includes(c)&&typeof v==='string'&&v.startsWith('=')?'':v)));
 put(total,['Tổng']);put(dt,['CHI TIẾT DỪNG XE GIỮA CHUYẾN TRÊN '+config.restMinutes+' PHÚT']);put(dt+1,['Chuyến ≥'+config.minKm+' km; gộp điểm dừng cùng địa chỉ hoặc ≤300 m, cách <10 phút, xe nhích <1 km.']);
 put(dh,['Chuyến','Thứ tự dừng','Địa chỉ','Bắt đầu dừng','Kết thúc dừng','Thời gian (phút)','Dòng đầu Data','Dòng cuối Data']);
 rests.forEach((r,i)=>put(df+i,[r.trip,r.order,TextUtils.safeSpreadsheetText(r.address),times[r.start],times[r.end],r.minutes,r.start+5,r.end+5]));if(!rests.length)put(df,['Không có lần dừng xe giữa chuyến trên '+config.restMinutes+' phút']);matrix[de-1][4]='Tổng';
 for(let i=0;i<matrix.length;i+=1000)out.getRangeByIndexes(i,0,Math.min(1000,matrix.length-i),15).values=matrix.slice(i,i+1000);
 trips.forEach((t,i)=>[1,2,3,6,7,10,11].forEach(c=>{if(typeof t[c]==='string'&&t[c].startsWith('='))out.getCell(i+4,c).formulas=[[t[c]]];}));
 if(trips.length){out.getRange('G'+total+':H'+total).formulas=[['=SUM(G5:G'+(total-1)+')','=SUM(H5:H'+(total-1)+')']];for(const a of ['E5:F'+(total-1),'M5:M'+(total-1)])out.getRange(a).numberFormat=Array.from({length:trips.length},()=>Array(a.startsWith('E')?2:1).fill('dd/mm/yyyy hh:mm:ss'));for(const a of ['G5:H'+(total-1),'K5:L'+(total-1)])out.getRange(a).numberFormat=Array.from({length:trips.length},()=>['0.00','0.00']);}else out.getRange('G'+total+':H'+total).values=[[0,0]];
 out.getRange('G'+total+':H'+total).numberFormat=[['0.00','0.00']];
 if(rests.length){out.getRange('D'+df+':E'+(de-1)).numberFormat=rests.map(()=>['dd/mm/yyyy hh:mm:ss','dd/mm/yyyy hh:mm:ss']);out.getRange('F'+df+':F'+(de-1)).numberFormat=rests.map(()=>['0.0']);out.getRange('F'+de).formulas=[['=SUM(F'+df+':F'+(de-1)+')']];}else out.getRange('F'+de).values=[[0]];
 out.getRange('F'+de).numberFormat=[['0.0']];
 for(const a of ['A1:O1','A'+dt+':H'+dt]){const f=out.getRange(a).format;f.fill.color='#214365';f.font.color='#FFFFFF';f.font.bold=true;}
 for(const a of ['A4:O4','A'+dh+':H'+dh]){const f=out.getRange(a).format;f.fill.color='#DCEAF5';f.font.bold=true;}
 for(const a of ['A'+total+':O'+total,'E'+de+':F'+de])out.getRange(a).format.font.bold=true;
 formatReport(out,{total,dt,dh,df,de,tripCount:trips.length,restCount:rests.length});
 out.activate();await ctx.sync();return 'GPS: Đã ghi '+trips.length+' chuyến ≥'+config.minKm+' km và '+rests.length+' lần dừng vào Chuyến đi 1.';
 });
}
return {run,calculate,init,save,clear,add,resetDefaults,toggleSettings,validateConfig,migrateLegacy,formatReport};
})();
