'use strict';
(() => {
  const $=id=>document.getElementById(id), money=n=>n===null?'—':new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',maximumFractionDigits:2}).format(n), pct=n=>n===null?'—':`${(n*100).toFixed(1)}%`;
  let model=null,filtered=[],page=0,version=0;
  const pageSize=25;
  function element(tag,value,className){const e=document.createElement(tag);if(value!==undefined)e.textContent=value;if(className)e.className=className;return e;}
  function option(value,label){const e=element('option',label);e.value=value;return e;}
  function options(id,values,all){const select=$(id),old=select.value;select.replaceChildren();if(all)select.append(option('',all));for(const v of values)select.append(option(v,v));if(values.includes(old))select.value=old;}
  function clearData(){model=null;filtered=[];page=0;$('dashboard').hidden=true;$('empty').hidden=false;for(const id of ['monthly','customers','vendor-rows','detail-rows','issue-rows','mapping','year','customer','vendor'])$(id).replaceChildren();for(const id of ['sales-value','profit-value','margin-value','rows-value'])$(id).textContent='—';for(const id of ['selection-summary','quality-note','quality-summary','table-count','rows-note'])$(id).textContent='';$('search').value='';$('basis').value='hw';$('quarter').value='';}
  function reset(){version++;clearData();$('workbook').value='';$('clear').disabled=true;$('filename').textContent='No workbook selected';$('file-details').textContent='Choose a file to start.';$('status').textContent='Your workbook stays local. No sales data is included in this site.';}
  $('clear').addEventListener('click',reset);window.addEventListener('pagehide',reset);
  $('workbook').addEventListener('click',()=>{$('workbook').value='';});
  $('workbook').addEventListener('change',async()=>{
    const file=$('workbook').files[0];if(!file)return;
    const current=++version;clearData();$('clear').disabled=false;$('filename').textContent=file.name;$('file-details').textContent=`${(file.size/1024).toFixed(1)} KB · Loaded ${new Date().toLocaleString()}`;$('status').textContent='Reading and validating workbook locally…';
    try{
      if(!/\.xlsx$/i.test(file.name)||!file.size||file.size>25*1024*1024)throw new Error('Choose a nonempty .xlsx workbook under 25 MB.');
      const buffer=await file.arrayBuffer();if(current!==version)return;
      const sig=new Uint8Array(buffer,0,Math.min(4,buffer.byteLength));if(sig[0]!==0x50||sig[1]!==0x4b)throw new Error('This file is not an .xlsx container.');
      if(!window.XLSX)throw new Error('The bundled Excel reader could not load. Keep the vendor folder beside this page.');
      DealLog.validateContainer(buffer);
      const wb=XLSX.read(buffer,{type:'array',cellFormula:true,cellHTML:false,cellText:false,cellDates:false});
      const parsed=DealLog.inspectWorkbook(wb,XLSX);if(current!==version)return;model=parsed;
      $('status').textContent=`Date Closed coverage: ${model.coverage[0]} to ${model.coverage.at(-1)}. ${model.records.length} valid log rows. This does not establish complete-year coverage.`;
      $('dashboard').hidden=false;$('empty').hidden=true;
      options('year',[...new Set(model.records.map(r=>r.date.slice(0,4)))].sort());$('year').value=[...new Set(model.records.map(r=>r.date.slice(0,4)))].sort().at(-1);
      showQuality();refresh(true);
    }catch(error){if(current!==version)return;clearData();$('status').textContent=`Could not use workbook: ${error.message||'Invalid Excel file.'}`;}
    finally{if(current===version)$('workbook').value='';}
  });
  function showQuality(){
    const r=model.records;$('quality-summary').textContent=`Source checks · ${model.issues.length} findings`;
    $('quality-note').textContent=`${model.sheets} workbook sheets; only the two current logs are included. ${r.filter(x=>!x.invoice).length}/${r.length} valid rows have no invoice number; ${r.filter(x=>!x.paid).length}/${r.length} have no Date Paid; ${r.filter(x=>!x.rep).length}/${r.length} have no REP. Missing accounting fields do not mean unpaid or unbilled. ${model.ignored} nontransaction/template rows ignored. ${model.recalculated} simple formula cells recalculated.`;
    $('issue-rows').replaceChildren();for(const issue of model.issues){const tr=element('tr');tr.append(element('td',issue.location),element('td',issue.message));$('issue-rows').append(tr);}
    if(!model.issues.length){const tr=element('tr');const td=element('td','No row-level exceptions found.');td.colSpan=2;tr.append(td);$('issue-rows').append(tr);}
    $('mapping').replaceChildren();for(const m of model.mapping)$('mapping').append(element('p',`${m.sheet} · Header row ${m.headerRow}. ${m.columns}`));
  }
  function baseRows(){return model.records.filter(r=>r.kind===($('basis').value==='hw'?'hw':'ps'));}
  function refresh(changeBasis=false){
    if(!model)return;page=0;const basis=$('basis').value,base=baseRows(),actual=basis==='actual';
    if(changeBasis){$('customer').value='';$('vendor').value='';options('customer',[...new Set(base.map(r=>r.customer))].sort(),'All customers');options('vendor',[...new Set(base.map(r=>r.vendor))].sort(),'All vendors');}
    const candidates=base.filter(r=>(!$('year').value||r.date.startsWith($('year').value))&&(!$('quarter').value||Math.ceil(Number(r.date.slice(5,7))/3)===Number($('quarter').value))&&(!$('customer').value||r.customer===$('customer').value)&&(!$('vendor').value||r.vendor===$('vendor').value));
    filtered=candidates.filter(r=>!actual||r.actualSold!==null).map(r=>({...r,sales:actual?r.actualSold:r.sold,profit:actual?r.actualProfit:r.gp}));
    const totals=DealLog.summarize(filtered),has=filtered.length>0;
    $('basis-note').textContent=basis==='hw'?'Hardware/software · Closed-log SOLD AT and GROSS PROFIT by Date Closed. These are not verified invoiced sales.':actual?'PS actuals · Actual SOLD AT and PROFIT, grouped by Date Closed. PROFIT is not labeled gross profit in the source.':'PS estimates · Estimated SOLD AT E and GROSS PROFIT E by Date Closed. Estimates are not actual or invoiced revenue.';
    const coverage=filtered.map(r=>r.date).sort();$('selection-summary').textContent=has?`${filtered.length} log rows · ${totals.customers} customers · Selected dates ${coverage[0]} to ${coverage.at(-1)}${actual?` · ${candidates.length-filtered.length} rows lack complete actuals`:''}`:(actual?'No complete PS actuals for these filters. Blank actuals are not zero revenue.':'No log rows match these filters. This does not establish zero activity.');
    const excluded=model.issues.filter(i=>i.message.includes('Excluded row')||i.message.includes('actuals excluded')).length;
    if(excluded)$('selection-summary').textContent+=` · ${excluded} source rows/actuals excluded: totals may be incomplete. Open Source checks.`;
    $('sales-label').textContent=basis==='ps'?'Estimated PS sales':actual?'Actual PS sales':'Closed-log sales';$('profit-label').textContent=basis==='ps'?'Estimated gross profit':actual?'Actual profit':'Gross profit';$('margin-label').textContent=actual?'Actual profit margin':'Gross margin';
    $('sales-value').textContent=has?money(totals.sold):'—';$('profit-value').textContent=has?money(totals.profit):'—';$('margin-value').textContent=has?pct(totals.margin):'—';$('rows-value').textContent=String(totals.rows);
    $('sales-note').textContent='Workbook $ units · Currency unverified';$('profit-note').textContent=actual?'Actuals PROFIT column':basis==='ps'?'GROSS PROFIT E column':'GROSS PROFIT column';$('rows-note').textContent=`${totals.customers} customers · Not a unique-deal count`;
    $('vendor-profit').textContent=actual?'Profit':'Gross profit';$('detail-profit').textContent=actual?'Profit':'Gross profit';renderMonthly();renderCustomers();renderVendors();renderTable();
  }
  function groups(key){const g=new Map();for(const r of filtered){const k=r[key];if(!g.has(k))g.set(k,[]);g.get(k).push(r);}return [...g].map(([label,rows])=>({label,...DealLog.summarize(rows)})).sort((a,b)=>b.sold-a.sold);}
  function barRow(label,value,max,note){const row=element('div',undefined,'bar-row'),head=element('div',undefined,'bar-head');head.append(element('span',label),element('strong',value===null?'No rows':money(value)));const meter=document.createElement('meter');meter.min=0;meter.max=max||1;meter.value=Math.abs(value||0);meter.setAttribute('aria-label',`${label}: ${value===null?'no rows':money(value)}`);if(value<0)meter.className='negative';row.append(head,meter);if(note)row.append(element('small',note));return row;}
  function renderMonthly(){const monthly=new Map();for(const r of filtered){const m=r.date.slice(5,7);monthly.set(m,(monthly.get(m)||0)+r.sales);}const max=Math.max(1,...[...monthly.values()].map(Math.abs));$('monthly').replaceChildren();const q=Number($('quarter').value);for(let m=1;m<=12;m++){if(q&&Math.ceil(m/3)!==q)continue;const key=String(m).padStart(2,'0'),label=new Date(Date.UTC(2026,m-1,1)).toLocaleString('en-US',{month:'short',timeZone:'UTC'});$('monthly').append(barRow(label,monthly.has(key)?monthly.get(key):null,max));}}
  function renderCustomers(){const rows=groups('customer'),max=Math.max(1,...rows.map(r=>Math.abs(r.sold)));$('customers').replaceChildren();if(!rows.length){$('customers').append(element('p','No matching rows.','small'));return;}for(const r of rows.slice(0,8))$('customers').append(barRow(r.label,r.sold,max,`${r.rows} log rows · ${pct(r.margin)} margin`));if(rows.length>8)$('customers').append(element('p',`Top 8 of ${rows.length} customers. Use the customer filter for others.`,'small'));}
  function td(value,numeric=false){return element('td',value,numeric?'numeric':undefined);}
  function renderVendors(){const rows=groups('vendor');$('vendor-rows').replaceChildren();for(const r of rows){const tr=element('tr');tr.append(td(r.label),td(r.rows,true),td(money(r.sold),true),td(money(r.profit),true),td(pct(r.margin),true));$('vendor-rows').append(tr);}if(!rows.length)emptyTable('vendor-rows',5);}
  function emptyTable(id,span){const tr=element('tr'),cell=td('No matching rows.');cell.colSpan=span;tr.append(cell);$(id).append(tr);}
  function renderTable(){const search=$('search').value.toLowerCase();const rows=filtered.filter(r=>[r.name,r.customer,r.vendor,r.id].some(v=>v.toLowerCase().includes(search))).sort((a,b)=>b.date.localeCompare(a.date)||a.row-b.row);const pages=Math.max(1,Math.ceil(rows.length/pageSize));page=Math.min(page,pages-1);$('detail-rows').replaceChildren();for(const r of rows.slice(page*pageSize,(page+1)*pageSize)){const tr=element('tr');tr.append(td(r.date),td(`${r.name||'(No name)'} · ${r.id||'(No ID)'}`),td(r.customer),td(r.vendor),td(money(r.sales),true),td(money(r.profit),true),td(pct(r.sales===0?null:r.profit/r.sales),true),td(`${r.sheet}!${r.row}`));if(r.profit<0)tr.className='loss';$('detail-rows').append(tr);}if(!rows.length)emptyTable('detail-rows',8);$('table-count').textContent=`${rows.length} rows · Page ${page+1} of ${pages}`;$('previous').disabled=page===0;$('next').disabled=page>=pages-1;}
  $('basis').addEventListener('change',()=>refresh(true));for(const id of ['year','quarter','customer','vendor'])$(id).addEventListener('change',()=>refresh());$('search').addEventListener('input',()=>{page=0;renderTable();});$('previous').addEventListener('click',()=>{page--;renderTable();});$('next').addEventListener('click',()=>{page++;renderTable();});
  $('reset-filters').addEventListener('click',()=>{$('quarter').value='';$('customer').value='';$('vendor').value='';$('search').value='';$('year').value=[...new Set(model.records.map(r=>r.date.slice(0,4)))].sort().at(-1);refresh();});
})();
