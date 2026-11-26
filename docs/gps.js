const GPS = (() => {
'use strict';
function calculate(data, config){
  config = validateConfig(config);
  const FIRST_DATA_ROW = 5;

  const MIN_FACTORY_STOP_MINUTES = 30;
  const MIN_REST_MINUTES = 10;
  const MIN_TRIP_KM = config.minKm;
  const MAX_MERGE_GAP_MINUTES = 10;
  const MAX_REST_MOVE_KM = 1;
  const MAX_SAME_PLACE_DISTANCE_KM = 0.3;

  const times = data.map(row => excelDate(row[0]));

  // 1. Nhận diện nhà máy bằng GPS, không dựa vào tên địa chỉ.
  const FACTORY_RADIUS_KM = config.radius;
  const factories = config.factories;
  const factoryCodes = data.map(row => {
    if (!validGps(row)) return "";
    for (const f of factories) {
      if (haversineKm(f.lat, f.lon, toNumber(row[2]),
        toNumber(row[1])) <= FACTORY_RADIUS_KM) return f.code;
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
  // Mốc đến: đầu lần dừng dài đầu tiên (>=30 phút), nếu có.
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
    if (!Number.isFinite(tripKm) || tripKm < MIN_TRIP_KM) {
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
        "không dừng đủ 30 phút.";
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

      if (minutes <= MIN_REST_MINUTES) continue;

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
    lon > 100 &&
    lon < 110 &&
    lat > 8 &&
    lat < 24
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
const KEY='office-utility.gps.settings.v1';
const fieldIds=['gps-nm1-lon','gps-nm1-lat','gps-nm2-lon','gps-nm2-lat','gps-radius','gps-min-km'];
function validateConfig(c){
 if(!c||!Array.isArray(c.factories)||c.factories.length!==2)throw Error('Nhập tọa độ NM1 và NM2.');
 const radius=c.radius,minKm=c.minKm;
 if(!Number.isFinite(radius)||radius<=0||radius>100)throw Error('Bán kính phải lớn hơn 0 và không vượt 100 km.');
 if(!Number.isFinite(minKm)||minKm<=0||minKm>100000)throw Error('Quãng đường tối thiểu phải lớn hơn 0 và không vượt 100.000 km.');
 const factories=c.factories.map((f,i)=>{
   if(!f||!Number.isFinite(f.lon)||!Number.isFinite(f.lat)||f.lon< -180||f.lon>180||f.lat< -90||f.lat>90)throw Error('Tọa độ NM'+(i+1)+' không hợp lệ.');
   return {code:'NM'+(i+1),lon:f.lon,lat:f.lat};
 });
 if(haversineKm(factories[0].lat,factories[0].lon,factories[1].lat,factories[1].lon)<=2*radius)throw Error('Hai vùng nhà máy bị chồng lấn. Kiểm tra tọa độ hoặc giảm bán kính.');
 return {factories,radius,minKm};
}
function readConfig(){
 const nums=fieldIds.map(id=>{const v=document.getElementById(id).value.trim().replace(',','.');return v===''?NaN:Number(v);});
 return validateConfig({factories:[{lon:nums[0],lat:nums[1]},{lon:nums[2],lat:nums[3]}],radius:nums[4],minKm:nums[5]});
}
function displayConfig(c){
 const values=c?[c.factories[0].lon,c.factories[0].lat,c.factories[1].lon,c.factories[1].lat,c.radius,c.minKm]:['','','','',1,5];
 fieldIds.forEach((id,i)=>document.getElementById(id).value=String(values[i]));
}
function init(){
 try{const v=localStorage.getItem(KEY);if(v)displayConfig(validateConfig(JSON.parse(v)));}
 catch(_){displayConfig(null);}
}
function save(){
 const c=readConfig();
 try{localStorage.setItem(KEY,JSON.stringify(c));}catch(_){throw Error('Không lưu được cục bộ. Bạn vẫn có thể nhập tọa độ và chạy trong phiên hiện tại.');}
 return 'Đã lưu cấu hình GPS trên thiết bị này.';
}
function clear(){
 try{localStorage.removeItem(KEY);}catch(_){throw Error('Không xóa được cấu hình trong bộ nhớ cục bộ.');}
 displayConfig(null);return 'Đã xóa cấu hình GPS trên thiết bị này.';
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
 put(2,['NM1 / NM2 | Bán kính GPS '+config.radius+' km']);put(3,['GPS – Chuyến ≥'+config.minKm+' km; dừng nhà máy ≥30 phút.']);
 put(4,['Chuyến','Điểm đi','Điểm xa nhất (điểm đến)','Điểm về','Bắt đầu','Kết thúc','Quãng đường (km)','Thời gian (giờ)','Dòng bắt đầu','Dòng kết thúc','Km đầu','Km cuối','Đến điểm xa nhất lúc','Dòng điểm xa nhất','Ghi chú']);
 trips.forEach((t,i)=>put(i+5,t.map((v,c)=>[1,2,3,6,7,10,11].includes(c)&&typeof v==='string'&&v.startsWith('=')?'':v)));
 put(total,['Tổng']);put(dt,['CHI TIẾT DỪNG NGHỈ TRÊN 10 PHÚT']);put(dt+1,['Chuyến ≥'+config.minKm+' km; gộp điểm dừng cùng địa chỉ hoặc ≤300 m, cách <10 phút, xe nhích <1 km.']);
 put(dh,['Chuyến','Thứ tự dừng','Địa chỉ','Bắt đầu dừng','Kết thúc dừng','Thời gian (phút)','Dòng đầu Data','Dòng cuối Data']);
 rests.forEach((r,i)=>put(df+i,[r.trip,r.order,TextUtils.safeSpreadsheetText(r.address),times[r.start],times[r.end],r.minutes,r.start+5,r.end+5]));if(!rests.length)put(df,['Không có điểm dừng trên 10 phút']);matrix[de-1][4]='Tổng';
 for(let i=0;i<matrix.length;i+=1000)out.getRangeByIndexes(i,0,Math.min(1000,matrix.length-i),15).values=matrix.slice(i,i+1000);
 trips.forEach((t,i)=>[1,2,3,6,7,10,11].forEach(c=>{if(typeof t[c]==='string'&&t[c].startsWith('='))out.getCell(i+4,c).formulas=[[t[c]]];}));
 if(trips.length){out.getRange('G'+total+':H'+total).formulas=[['=SUM(G5:G'+(total-1)+')','=SUM(H5:H'+(total-1)+')']];for(const a of ['E5:F'+(total-1),'M5:M'+(total-1)])out.getRange(a).numberFormat=Array.from({length:trips.length},()=>Array(a.startsWith('E')?2:1).fill('dd/mm/yyyy hh:mm:ss'));for(const a of ['G5:H'+(total-1),'K5:L'+(total-1)])out.getRange(a).numberFormat=Array.from({length:trips.length},()=>['0.00','0.00']);}else out.getRange('G'+total+':H'+total).values=[[0,0]];
 out.getRange('G'+total+':H'+total).numberFormat=[['0.00','0.00']];
 if(rests.length){out.getRange('D'+df+':E'+(de-1)).numberFormat=rests.map(()=>['dd/mm/yyyy hh:mm:ss','dd/mm/yyyy hh:mm:ss']);out.getRange('F'+df+':F'+(de-1)).numberFormat=rests.map(()=>['0.0']);out.getRange('F'+de).formulas=[['=SUM(F'+df+':F'+(de-1)+')']];}else out.getRange('F'+de).values=[[0]];
 out.getRange('F'+de).numberFormat=[['0.0']];
 for(const a of ['A1:O1','A'+dt+':H'+dt]){const f=out.getRange(a).format;f.fill.color='#214365';f.font.color='#FFFFFF';f.font.bold=true;}
 for(const a of ['A4:O4','A'+dh+':H'+dh]){const f=out.getRange(a).format;f.fill.color='#DCEAF5';f.font.bold=true;}
 for(const a of ['A'+total+':O'+total,'E'+de+':F'+de])out.getRange(a).format.font.bold=true;
 [100,180,290,160,160,160,110,110,100,100,100,100,160,110,300].forEach((w,i)=>out.getRangeByIndexes(0,i,1,1).format.columnWidth=w);
 for(const a of ['A4:O4','A'+dh+':H'+dh]){const f=out.getRange(a).format;f.wrapText=true;f.verticalAlignment='Center';f.autofitRows();}
 if(trips.length)out.getRange('O5:O'+(total-1)).format.wrapText=true;
 out.activate();await ctx.sync();return 'GPS: Đã ghi '+trips.length+' chuyến ≥'+config.minKm+' km và '+rests.length+' lần dừng vào Chuyến đi 1.';
 });
}
return {run,calculate,init,save,clear,validateConfig};
})();
