/* Source-specific validation and arithmetic. No spreadsheet formula is executed as JavaScript. */
(function (root) {
  'use strict';
  const common = ['Deal ID','Date Closed','COMPANY','Deal Name','Customer','Vendor PO#','Vendor','Deduction','REP','ADDITIONAL NOTES/COMMENTS','DATE PAID','INVOICE #','MONTH CLOSED','Quarter Included'];
  const schemas = [
    {name:'2026 HW+SW Log',kind:'hw',headers:[...common,'Customer PO#','BOUGHT AT','SOLD AT','PROFIT','GROSS PROFIT']},
    {name:'2026 PS Log',kind:'ps',headers:[...common,'Customer PO#/SOW','BOUGHT AT E','SOLD AT E','PROFIT E','GROSS PROFIT E','BOUGHT AT','SOLD AT','PROFIT','Project Number']}
  ];
  const norm = v => String(v ?? '').trim().replace(/\s+/g,' ').toUpperCase();
  const text = v => String(v ?? '').trim();
  const present = v => v !== undefined && v !== null && v !== '';
  function parseDate(value, date1904, XLSX) {
    let y,m,d;
    if (typeof value === 'number') {
      const p=XLSX.SSF.parse_date_code(value,{date1904});
      if (!p) return null;
      [y,m,d]=[p.y,p.m,p.d];
    } else {
      const s=text(value); let match;
      if ((match=/^(\d{4})-(\d{2})-(\d{2})(?:T00:00:00(?:\.000)?Z?)?$/.exec(s))) [,y,m,d]=match.map(Number);
      else if ((match=/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(s))) {m=+match[1];d=+match[2];y=+match[3];if(y<100)y+=2000;}
      else return null;
    }
    const check=new Date(Date.UTC(y,m-1,d));
    if(y<1900||y>2200||check.getUTCFullYear()!==y||check.getUTCMonth()!==m-1||check.getUTCDate()!==d) return null;
    return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
  }
  function arithmetic(formula, resolve) {
    const source=formula.replace(/^=/,'').replace(/\s+/g,'').toUpperCase();
    const tokens=source.match(/\$?[A-Z]{1,3}\$?\d+|(?:\d+(?:\.\d*)?|\.\d+)|[()+\-*/]/g)||[];
    if(tokens.join('')!==source||tokens.length>120) throw new Error('Unsupported formula. Save calculated values or use simple arithmetic.');
    let i=0;
    function atom(){const t=tokens[i++];if(t==='-')return -atom();if(t==='+')return atom();if(t==='('){const v=add();if(tokens[i++]!==')')throw new Error('Invalid formula.');return v;}if(/^\$?[A-Z]/.test(t||''))return resolve(t.replace(/\$/g,''));if(t&&/^(\d|\.)/.test(t))return Number(t);throw new Error('Invalid formula.');}
    function mul(){let v=atom();while(tokens[i]==='*'||tokens[i]==='/'){const op=tokens[i++],b=atom();if(op==='/'&&b===0)throw new Error('Division by zero.');v=op==='*'?v*b:v/b;}return v;}
    function add(){let v=mul();while(tokens[i]==='+'||tokens[i]==='-'){const op=tokens[i++],b=mul();v=op==='+'?v+b:v-b;}return v;}
    const result=add();if(i!==tokens.length||!Number.isFinite(result))throw new Error('Invalid formula result.');return result;
  }
  function inspectWorkbook(wb,XLSX){
    const records=[],issues=[],mapping=[];let ignored=0,recalculated=0;
    const date1904=!!wb.Workbook?.WBProps?.date1904;
    for(const spec of schemas){
      const sheet=wb.Sheets[spec.name];
      if(!sheet)throw new Error(`Missing sheet “${spec.name}”. Keep both log tabs or update the mapping before using this workbook.`);
      const range=XLSX.utils.decode_range(sheet['!ref']||'A1');
      if(range.e.r>20000||range.e.c>200)throw new Error(`${spec.name}: sheet is too large (maximum 20,000 rows and 201 columns).`);
      let headerRow=-1,columns;
      for(let r=0;r<Math.min(12,range.e.r+1);r++){
        const entries=[];for(let c=0;c<=range.e.c;c++)entries.push([norm(sheet[XLSX.utils.encode_cell({r,c})]?.v),c]);
        if(entries.some(([v])=>v==='DEAL ID')&&entries.some(([v])=>v==='DATE CLOSED')){
          headerRow=r;columns=new Map(entries.filter(([v])=>v));
          const duplicates=entries.filter(([v])=>v&&entries.filter(([x])=>x===v).length>1);
          if(duplicates.length)throw new Error(`${spec.name}: duplicate column headers. Resolve them before importing.`);
          break;
        }
      }
      if(headerRow<0)throw new Error(`${spec.name}: cannot find the Deal ID / Date Closed header row.`);
      const missing=spec.headers.filter(h=>!columns.has(norm(h)));
      if(missing.length)throw new Error(`${spec.name}: missing or renamed columns: ${missing.join(', ')}. Restore these headers or update the mapping; metrics have been cleared.`);
      mapping.push({sheet:spec.name,headerRow:headerRow+1,columns:spec.headers.map(h=>`${h}: ${XLSX.utils.encode_col(columns.get(norm(h)))}`).join(' · ')});
      const cell=(r,h)=>sheet[XLSX.utils.encode_cell({r,c:columns.get(norm(h))})];
      const val=(r,h)=>cell(r,h)?.v;
      for(let r=headerRow+1;r<=range.e.r;r++){
        const soldLabel=spec.kind==='ps'?'SOLD AT E':'SOLD AT';
        // IDs and formula-only template rows are not transactions. Summary formulas have no business identity.
        const hasIdentity=['Date Closed','Deal Name','Customer','Vendor'].some(h=>present(val(r,h)));
        const rawSold=cell(r,soldLabel);
        if(!hasIdentity&&!(rawSold&&!rawSold.f&&typeof rawSold.v==='number')){ignored++;continue;}
        if(norm(val(r,soldLabel))===norm(soldLabel)&&!hasIdentity){ignored++;continue;}
        const row=r+1, location=`${spec.name}!${row}`;
        const warn=message=>issues.push({location,message});
        const date=parseDate(val(r,'Date Closed'),date1904,XLSX);
        if(!date){warn('Excluded row: missing or invalid Date Closed (text dates use MM/DD/YY).');continue;}
        const values=new Map();
        function numericAddress(address,stack=[],allowBlank=false){
          if(values.has(address))return values.get(address);
          const decoded=XLSX.utils.decode_cell(address);
          if(decoded.r!==r)throw new Error('Cross-row formula requires review.');
          if(stack.includes(address)||stack.length>12)throw new Error('Circular or excessively nested formula.');
          const c=sheet[address];
          if(!c||(!present(c.v)&&!c.f)){if(allowBlank)return 0;throw new Error(`Missing number in ${address}.`);}
          if(c.t==='e')throw new Error(`Excel error in ${address}.`);
          let n;
          if(c.f){
            n=arithmetic(c.f,ref=>numericAddress(ref,[...stack,address],ref===XLSX.utils.encode_cell({r,c:columns.get('DEDUCTION')})));
            recalculated++;
            if(typeof c.v==='number'&&Math.abs(n-c.v)>0.011)warn(`${address}: formula cache differs; recalculated simple arithmetic is used.`);
          }else{if(typeof c.v!=='number'||!Number.isFinite(c.v))throw new Error(`Expected a numeric value in ${address}.`);n=c.v;}
          values.set(address,n);return n;
        }
        const number=h=>numericAddress(XLSX.utils.encode_cell({r,c:columns.get(norm(h))}));
        let sold,cost,gp,actualSold=null,actualCost=null,actualProfit=null;
        try{
          sold=number(soldLabel);cost=number(spec.kind==='ps'?'BOUGHT AT E':'BOUGHT AT');gp=number(spec.kind==='ps'?'GROSS PROFIT E':'GROSS PROFIT');
        }catch(e){warn(`Excluded row: ${e.message}`);continue;}
        if(spec.kind==='ps'&&['BOUGHT AT','SOLD AT','PROFIT'].some(h=>present(val(r,h))||cell(r,h)?.f)){
          try{actualSold=number('SOLD AT');actualCost=number('BOUGHT AT');actualProfit=number('PROFIT');}
          catch(e){actualSold=actualCost=actualProfit=null;warn(`PS actuals excluded: ${e.message}`);}
        }
        const customer=text(val(r,'Customer'))||'(Not provided)',vendor=text(val(r,'Vendor'))||'(Not provided)';
        if(!present(val(r,'Customer')))warn('Customer missing; shown as Not provided.');
        if(!present(val(r,'Vendor')))warn('Vendor missing; no OEM is inferred.');
        if(gp<0)warn('Negative gross profit retained in totals.');
        records.push({kind:spec.kind,row,sheet:spec.name,id:text(val(r,'Deal ID')),date,customer,vendor,name:text(val(r,'Deal Name')),sold,cost,gp,actualSold,actualCost,actualProfit,invoice:present(val(r,'INVOICE #')),paid:present(val(r,'DATE PAID')),rep:present(val(r,'REP'))});
      }
    }
    if(!records.length)throw new Error('No valid dated sales rows found. '+issues.slice(0,3).map(i=>`${i.location}: ${i.message}`).join(' '));
    const counts=new Map();for(const r of records){if(r.id){const k=`${r.sheet}|${r.id}`;counts.set(k,(counts.get(k)||0)+1);}}
    for(const [key,count] of counts){if(count>1)issues.push({location:key.replace('|',' · ID '),message:`Repeated ID on ${count} rows. All rows retained; counts represent log rows, not unique deals.`});}
    return {records,issues,mapping,ignored,recalculated,sheets:wb.SheetNames.length,coverage:[...records.map(r=>r.date)].sort().filter((v,i,a)=>i===0||i===a.length-1)};
  }
  function summarize(records){const sold=records.reduce((s,r)=>s+r.sales,0),profit=records.reduce((s,r)=>s+r.profit,0);return {sold,profit,margin:sold===0?null:profit/sold,rows:records.length,customers:new Set(records.map(r=>r.customer)).size};}
  function validateContainer(buffer){
    const bytes=buffer instanceof Uint8Array?buffer:new Uint8Array(buffer),view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
    if(bytes.length<22||bytes.length>25*1024*1024)throw new Error('Invalid workbook size.');
    let end=-1;
    for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--){if(view.getUint32(i,true)===0x06054b50&&i+22+view.getUint16(i+20,true)===bytes.length){end=i;break;}}
    if(end<0)throw new Error('The workbook ZIP container is incomplete or unsupported.');
    const count=view.getUint16(end+10,true);let offset=view.getUint32(end+16,true),total=0;
    if(count>5000||view.getUint16(end+4,true)!==0||view.getUint16(end+6,true)!==0)throw new Error('Unsupported multipart or oversized workbook.');
    for(let n=0;n<count;n++){
      if(offset+46>end||view.getUint32(offset,true)!==0x02014b50)throw new Error('Invalid workbook directory.');
      const size=view.getUint32(offset+24,true),flags=view.getUint16(offset+8,true);
      if(flags&1||size>30*1024*1024||(total+=size)>80*1024*1024)throw new Error('Encrypted or excessively expanded workbook is not supported.');
      offset+=46+view.getUint16(offset+28,true)+view.getUint16(offset+30,true)+view.getUint16(offset+32,true);
      if(offset>end)throw new Error('Invalid workbook directory length.');
    }
  }
  const api={inspectWorkbook,parseDate,arithmetic,summarize,validateContainer};
  root.DealLog=api;if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(globalThis);
