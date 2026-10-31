const picker=document.getElementById('picker');
const tabs=document.getElementById('tabs');
const area=document.getElementById('tableArea');
const title=document.getElementById('sheetTitle');
const search=document.getElementById('search');
const layoutToggle=document.getElementById('layoutToggle');
const editToggle=document.getElementById('editToggle');
const saveButton=document.getElementById('saveChanges');
const cancelButton=document.getElementById('cancelChanges');
const notice=document.getElementById('notice');
const CLIENT_KEY='diakinisi-ms-client-id';
const SHARE_KEY='diakinisi-onedrive-share-url';
const DEFAULT_CLIENT_ID='ca63a02f-82fb-4ff7-ac17-143c2ae76c9c';
const GRAPH='https://graph.microsoft.com/v1.0';
const MSAL_REDIRECT=new URL('auth.html',location.href).href;
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RELNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKGNS='http://schemas.openxmlformats.org/package/2006/relationships';
const descriptions={
  'ΠΑΡΑΓΓΕΛΙΕΣ':'Παραγγελίες και κατάσταση ολοκλήρωσης',
  'ΠΡΑΚΤΟΡΕΙΑ-ΠΕΛΑΤΕΣ-ΣΥΝΕΡΓΑΤΕΣ':'Αποστολές προς πρακτορεία και συνεργάτες',
  'ΠΡΟΣ ΣΥΝΕΡΓΕΙΑ':'Εργασίες και υλικά που δίνονται σε συνεργεία',
  'ΑΠΟ ΣΥΝΕΡΓΕΙΑ':'Παραλαβές και σχόλια από συνεργεία',
  'ΔΡΟΜΟΛΟΓΙΟ':'Πρόγραμμα διαδρομών και παραδόσεων'
};
let sheets=[];
let active=0;
let msalClient=null;
let msalAccount=null;
let remoteFile=null;
let originalBytes=null;
let editMode=false;
let dirty=new Map();
let styleDefinitions=[];
let cellStyleIds=new Map();
let themeColors=[];
let mobileCards=true;

document.getElementById('upload').onclick=()=>picker.click();
picker.onchange=async event=>{
  const file=event.target.files?.[0]; if(!file)return;
  try{
    remoteFile=null; originalBytes=null; dirty.clear(); editMode=false;
    await loadBytes(new Uint8Array(await file.arrayBuffer()),file.name);
    setNotice('Το αρχείο φορτώθηκε μόνο για προβολή. Για αποθήκευση αλλαγών, σύνδεσε το OneDrive.');
  }catch(error){fail(error)}
};
document.getElementById('oneDrive').onclick=connectOneDrive;
editToggle.onclick=()=>{if(editMode&&dirty.size){alert('Αποθήκευσε ή ακύρωσε πρώτα τις αλλαγές.');return}editMode=!editMode;renderControls();showSheet(active)};
cancelButton.onclick=()=>{dirty.clear();editMode=false;renderControls();showSheet(active)};
saveButton.onclick=saveChanges;
search.oninput=()=>{if(sheets[active])showSheet(active)};
layoutToggle.onclick=()=>{mobileCards=!mobileCards;showSheet(active)};

async function connectOneDrive(){
  if(!location.protocol.startsWith('http')){alert('Για σύνδεση OneDrive, η εφαρμογή πρέπει πρώτα να δημοσιευτεί σε ασφαλή διεύθυνση HTTPS.');return}
  let clientId=localStorage.getItem(CLIENT_KEY)||DEFAULT_CLIENT_ID;
  let shareUrl=localStorage.getItem(SHARE_KEY)||'';
  if(!clientId)return;localStorage.setItem(CLIENT_KEY,clientId.trim());
  if(!shareUrl){shareUrl=prompt('Επικόλλησε τον σύνδεσμο κοινής χρήσης του Excel. Θα αποθηκευτεί μόνο σε αυτή τη συσκευή.', '')||'';if(!shareUrl)return;localStorage.setItem(SHARE_KEY,shareUrl.trim())}
  try{
    setNotice('Γίνεται σύνδεση στον Microsoft λογαριασμό σου…');
    const token=await getAccessToken(clientId);
    const shareToken=encodeShareUrl(shareUrl);
    const metadata=await graphJson(`${GRAPH}/shares/${encodeURIComponent(shareToken)}/driveItem?$select=id,name,eTag,parentReference,webUrl`,token);
    const driveId=metadata.parentReference?.driveId;
    if(!driveId||!metadata.id)throw new Error('Δεν βρέθηκαν τα στοιχεία του κοινόχρηστου αρχείου. Έλεγξε τον σύνδεσμο και τη σύνδεση Microsoft.');
    const fileResponse=await graphFetch(`${GRAPH}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(metadata.id)}/content`,token);
    if(!fileResponse.ok)throw new Error(`Η λήψη του αρχείου απέτυχε (${fileResponse.status}). Έλεγξε ότι έχεις πρόσβαση στο κοινόχρηστο αρχείο.`);
    originalBytes=new Uint8Array(await fileResponse.arrayBuffer());
    remoteFile={id:metadata.id,driveId,name:metadata.name,eTag:metadata.eTag,shareToken};
    await loadBytes(originalBytes,metadata.name);
    editToggle.hidden=false;
    setNotice('Συνδέθηκες στο OneDrive. Μπορείς να επεξεργαστείς τα κελιά και να αποθηκεύσεις στο κοινόχρηστο αρχείο.');
  }catch(error){fail(error)}
}

async function getAccessToken(clientId){
  if(!window.msal?.PublicClientApplication)throw new Error('Δεν φορτώθηκε η βιβλιοθήκη σύνδεσης Microsoft. Άνοιξε ξανά την εφαρμογή με σύνδεση στο διαδίκτυο.');
  if(!msalClient||msalClient.config.auth.clientId!==clientId){
    msalClient=new msal.PublicClientApplication({auth:{clientId,authority:'https://login.microsoftonline.com/consumers',redirectUri:MSAL_REDIRECT},cache:{cacheLocation:'sessionStorage'}});
  }
  const accounts=msalClient.getAllAccounts();
  msalAccount=accounts[0]||null;
  if(!msalAccount){const login=await msalClient.loginPopup({scopes:['Files.ReadWrite'],redirectUri:MSAL_REDIRECT});msalAccount=login.account}
  try{return (await msalClient.acquireTokenSilent({scopes:['Files.ReadWrite'],account:msalAccount})).accessToken}
  catch{return (await msalClient.acquireTokenPopup({scopes:['Files.ReadWrite'],account:msalAccount,redirectUri:MSAL_REDIRECT})).accessToken}
}

function encodeShareUrl(url){
  const bytes=new TextEncoder().encode(url.trim());let binary='';
  for(let i=0;i<bytes.length;i+=0x8000)binary+=String.fromCharCode(...bytes.subarray(i,i+0x8000));
  return 'u!'+btoa(binary).replace(/=/g,'').replace(/\+/g,'-').replace(/\//g,'_');
}
async function graphFetch(url,token,options={}){return fetch(url,{...options,headers:{Authorization:`Bearer ${token}`,...options.headers}})}
async function graphJson(url,token,options={}){const response=await graphFetch(url,token,options);if(!response.ok)throw new Error(`Η Microsoft απάντησε με σφάλμα ${response.status}: ${await response.text()}`);return response.json()}

async function loadBytes(bytes,name){
  const workbook=XLSX.read(bytes,{type:'array',cellDates:false,cellStyles:true,cellHTML:true});
  const presentation=await readPresentation(bytes,workbook);
  styleDefinitions=presentation.styles;cellStyleIds=presentation.cellStyles;themeColors=presentation.themeColors;
  sheets=workbook.SheetNames.map(sheetName=>({name:sheetName,worksheet:workbook.Sheets[sheetName],rows:XLSX.utils.sheet_to_json(workbook.Sheets[sheetName],{header:1,defval:'',blankrows:true}),formulaCells:findFormulaCells(workbook.Sheets[sheetName])}));
  document.getElementById('filename').textContent=name;
  document.getElementById('filehint').textContent=`${sheets.length} φύλλα εργασίας · ${Math.round(bytes.byteLength/1024)} KB`;
  document.getElementById('status').textContent=remoteFile?'Συνδεδεμένο με OneDrive':'Αρχείο φορτώθηκε';
  document.querySelector('.dot').style.background='#38a16c';
  document.getElementById('sheetCount').textContent=`${sheets.length} ${sheets.length===1?'φύλλο':'φύλλα'}`;
  search.hidden=false;renderTabs();showSheet(0,true);renderControls();
}
function findFormulaCells(sheet){const result=new Set();for(const [address,cell] of Object.entries(sheet)){if(address[0]==='!')continue;if(cell?.f)result.add(address)}return result}
async function readPresentation(bytes,workbook){
  let zip;try{zip=await JSZip.loadAsync(bytes)}catch{return{styles:[],cellStyles:new Map(),themeColors:[]}}
  const themeColors=await readThemeColors(zip);const stylesFile=zip.file('xl/styles.xml');const styles=stylesFile?parseCellStyles(await stylesFile.async('string'),themeColors):[];const cellStyles=new Map();
  const paths=await getSheetPaths(zip);
  for(const [name,path] of paths){const file=zip.file(path);if(!file)continue;const doc=xmlDoc(await file.async('string'));const ids=new Map();for(const cell of doc.getElementsByTagNameNS(NS,'c')){const id=cell.getAttribute('s');if(id)ids.set(cell.getAttribute('r'),Number(id))}cellStyles.set(name,ids)}
  return{styles,cellStyles,themeColors};
}
async function readThemeColors(zip){
  const file=zip.file('xl/theme/theme1.xml');if(!file)return[];const doc=xmlDoc(await file.async('string'));const dml='http://schemas.openxmlformats.org/drawingml/2006/main';const scheme=doc.getElementsByTagNameNS(dml,'clrScheme')[0];if(!scheme)return[];
  const byName=new Map([...scheme.children].map(node=>[node.localName,node]));const order=['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
  return order.map(name=>{const color=byName.get(name)?.firstElementChild;if(!color)return'';return color.getAttribute('lastClr')||color.getAttribute('val')||''});
}
function parseCellStyles(source,themes){
  const doc=xmlDoc(source);const fonts=[...doc.getElementsByTagNameNS(NS,'fonts')[0]?.children||[]].map(node=>({name:childEl(node,'name')?.getAttribute('val'),size:Number(childEl(node,'sz')?.getAttribute('val'))||0,bold:!!childEl(node,'b'),italic:!!childEl(node,'i'),strike:!!childEl(node,'strike'),underline:!!childEl(node,'u'),color:excelColor(childEl(node,'color'),themes)}));
  const fills=[...doc.getElementsByTagNameNS(NS,'fills')[0]?.children||[]].map(node=>{const pattern=childEl(node,'patternFill');return{pattern:pattern?.getAttribute('patternType'),color:excelColor(childEl(pattern,'fgColor'),themes)}});
  const borders=[...doc.getElementsByTagNameNS(NS,'borders')[0]?.children||[]].map(node=>{const sides={};for(const side of ['left','right','top','bottom']){const edge=childEl(node,side);sides[side]={style:edge?.getAttribute('style')||'',color:excelColor(childEl(edge,'color'),themes)}}return sides});
  const xfs=doc.getElementsByTagNameNS(NS,'cellXfs')[0];if(!xfs)return[];
  return[...xfs.children].map(xf=>{const alignment=childEl(xf,'alignment');return{font:fonts[Number(xf.getAttribute('fontId'))]||{},fill:fills[Number(xf.getAttribute('fillId'))]||{},border:borders[Number(xf.getAttribute('borderId'))]||{},alignment:alignment?Object.fromEntries([...alignment.attributes].map(a=>[a.name,a.value])):{}}});
}
function childEl(node,name){return node?[...node.children].find(child=>child.localName===name):null}
function excelColor(node,themes){if(!node)return'';const rgb=node.getAttribute('rgb');if(rgb&&/^[0-9A-Fa-f]{6,8}$/.test(rgb))return`#${rgb.slice(-6)}`;const theme=node.getAttribute('theme');if(theme!==null){const color=themes[Number(theme)];return color&&/^[0-9A-Fa-f]{6}$/.test(color)?`#${color}`:''}return''}
function styleCss(style){
  const css=[];const font=style.font||{};const fill=style.fill||{};const align=style.alignment||{};const border=style.border||{};
  if(font.color)css.push(`color:${font.color}`);if(font.bold)css.push('font-weight:700');if(font.italic)css.push('font-style:italic');if(font.strike)css.push('text-decoration:line-through');else if(font.underline)css.push('text-decoration:underline');
  if(font.name){const family=font.name.replace(/[^\w ,.-]/g,'');if(family)css.push(`font-family:${family}`)}if(font.size)css.push(`font-size:${Math.max(7,Math.min(48,font.size))}pt`);
  if(fill.pattern==='solid'&&fill.color)css.push(`background-color:${fill.color}`);
  const horizontal={left:'left',center:'center',right:'right',justify:'justify',distributed:'justify',centerContinuous:'center'}[align.horizontal];if(horizontal)css.push(`text-align:${horizontal}`);
  const vertical={top:'top',center:'middle',bottom:'bottom',justify:'middle',distributed:'middle'}[align.vertical];if(vertical)css.push(`vertical-align:${vertical}`);
  if(align.wrapText==='1')css.push('white-space:pre-wrap');if(align.textRotation==='90')css.push('writing-mode:vertical-rl');if(Number(align.indent)>0)css.push(`padding-left:${Math.min(32,Number(align.indent)*8)}px`);
  for(const side of ['top','right','bottom','left']){const edge=border[side];const cssBorder=excelBorder(edge);if(cssBorder)css.push(`border-${side}:${cssBorder}`)}
  return css.join(';');
}
function excelBorder(edge){if(!edge?.style)return'';const map={hair:'1px solid',thin:'1px solid',dotted:'1px dotted',dashed:'1px dashed',dashDot:'1px dashed',dashDotDot:'1px dashed',medium:'2px solid',mediumDashed:'2px dashed',mediumDashDot:'2px dashed',mediumDashDotDot:'2px dashed',thick:'3px solid',double:'3px double'};const result=map[edge.style];return result?`${result}${edge.color?` ${edge.color}`:''}`:''}
function hyperlinkMarkup(link,text){const target=String(link.Target||'').trim();if(!/^https?:\/\/|^mailto:/i.test(target))return esc(text);return`<a href="${esc(target)}" rel="noopener noreferrer" target="_blank"${link.Tooltip?` title="${esc(link.Tooltip)}"`:''}>${esc(text)}</a>`}
function renderTabs(){tabs.innerHTML='<div class="sidehead">Φύλλα εργασίας</div>'+sheets.map((sheet,index)=>`<button class="cat ${index===active?'active':''}" data-i="${index}" title="${esc(descriptions[sheet.name]||sheet.name)}"><span class="mini">▤</span>${esc(sheet.name)}<span class="badge">${Math.max(0,sheet.rows.length-1)}</span></button>`).join('');tabs.querySelectorAll('.cat').forEach(button=>button.onclick=()=>showSheet(+button.dataset.i,true))}
function showSheet(index,scrollToEnd=false){
  active=index;renderTabs();const sheet=sheets[index];if(!sheet)return;title.textContent=sheet.name;
  const ws=sheet.worksheet;const ref=ws['!ref']||'A1:A1';const range=XLSX.utils.decode_range(ref);const query=search.value.trim().toLocaleLowerCase('el');
  const merges=ws['!merges']||[];const anchors=new Map();const covered=new Set();
  for(const merge of merges){anchors.set(`${merge.s.r},${merge.s.c}`,merge);for(let r=merge.s.r;r<=merge.e.r;r++)for(let c=merge.s.c;c<=merge.e.c;c++)if(r!==merge.s.r||c!==merge.s.c)covered.add(`${r},${c}`)}
  const mobile=window.matchMedia('(max-width:740px)').matches;
  layoutToggle.hidden=!mobile;layoutToggle.textContent=mobileCards?'▦ Προβολή Excel':'▤ Κάρτες';
  if(mobile&&mobileCards){renderMobileCards(sheet,range,query,anchors,covered,scrollToEnd);return}
  const rowHasMatch=new Set();
  if(query)for(let r=range.s.r;r<=range.e.r;r++)for(let c=range.s.c;c<=range.e.c;c++){const cell=ws[XLSX.utils.encode_cell({r,c})];if(String(cell?.w??cell?.v??'').toLocaleLowerCase('el').includes(query))rowHasMatch.add(r)}
  const colNodes=[];for(let c=range.s.c;c<=range.e.c;c++){const column=ws['!cols']?.[c];const width=column?.wpx|| (column?.wch?column.wch*7+5:112);colNodes.push(`<col style="width:${Math.max(28,Math.min(500,width))}px">`)}
  const rows=[];
  for(let r=range.s.r;r<=range.e.r;r++){
    const rowMeta=ws['!rows']?.[r];if(rowMeta?.hidden)continue;
    const rowHeight=rowMeta?.hpx?` style="height:${rowMeta.hpx}px"`:rowMeta?.hpt?` style="height:${rowMeta.hpt*4/3}px"`:'';
    const cells=[];
    for(let c=range.s.c;c<=range.e.c;c++){
      if(covered.has(`${r},${c}`))continue;
      const address=XLSX.utils.encode_cell({r,c});const key=`${sheet.name}!${address}`;const cell=ws[address]||{};const merge=anchors.get(`${r},${c}`);const span=merge?` colspan="${merge.e.c-merge.s.c+1}" rowspan="${merge.e.r-merge.s.r+1}"`:'';
      const styleId=cellStyleIds.get(sheet.name)?.get(address)??0;const style=styleDefinitions[styleId]||styleDefinitions[0]||{};const locked=sheet.formulaCells.has(address);const editable=editMode&&!locked;const changed=dirty.has(key);
      const raw=dirty.get(key)?.value??cell.w??cell.v??'';const headerCell=ws[XLSX.utils.encode_cell({r:range.s.r,c})];const header=String(headerCell?.w??headerCell?.v??'');const display=formatCellValue(cell,raw,header);
      const valueMarkup=cell.l?.Target?hyperlinkMarkup(cell.l,display):esc(display);
      const hit=query&&rowHasMatch.has(r)?' search-match':'';
      const titleText=locked?'Τύπος Excel — δεν επεξεργάζεται από την εφαρμογή':cell.l?.Tooltip|| (editable?'Πάτησε για επεξεργασία':'');
      cells.push(`<td class="edit-cell ${locked?'locked ':''}${changed?'changed':''}${hit}" ${span} ${editable?'contenteditable="true" spellcheck="false"':''} data-address="${address}" data-sheet="${esc(sheet.name)}" data-header="${esc(header)}" title="${esc(titleText)}" style="${styleCss(style)}">${valueMarkup}</td>`);
    }
    rows.push(`<tr${rowHeight}>${cells.join('')}</tr>`);
  }
  area.className='tablewrap excel-grid'+(editMode?' editing':'');
  area.innerHTML=`<table><colgroup>${colNodes.join('')}</colgroup><tbody>${rows.join('')}</tbody></table>`;
  area.querySelectorAll('[contenteditable="true"]').forEach(cell=>cell.addEventListener('input',()=>{const key=`${cell.dataset.sheet}!${cell.dataset.address}`;dirty.set(key,{sheet:cell.dataset.sheet,address:cell.dataset.address,value:cell.textContent,header:cell.dataset.header});cell.classList.add('changed');saveButton.disabled=false}));
  area.querySelectorAll('a').forEach(link=>link.addEventListener('click',event=>{if(editMode)event.preventDefault()}));
  if(query&&rowHasMatch.size)area.querySelector('.search-match')?.scrollIntoView({block:'nearest',inline:'nearest'});
  else if(scrollToEnd)requestAnimationFrame(()=>area.lastElementChild?.scrollIntoView({block:'end'}));
}
function renderMobileCards(sheet,range,query,anchors,covered,scrollToEnd=false){
  const ws=sheet.worksheet;const headerRow=range.s.r;const records=[];let currentGroup='';let currentGroupStyle='';let activeGroup='';
  const groupColumn=sheet.name==='ΠΑΡΑΓΓΕΛΙΕΣ'?1:['ΠΡΑΚΤΟΡΕΙΑ-ΠΕΛΑΤΕΣ-ΣΥΝΕΡΓΑΤΕΣ','ΠΡΟΣ ΣΥΝΕΡΓΕΙΑ','ΑΠΟ ΣΥΝΕΡΓΕΙΑ'].includes(sheet.name)?0:null;
  const isRoute=sheet.name==='ΔΡΟΜΟΛΟΓΙΟ';
  if(isRoute){const routeCell=ws[XLSX.utils.encode_cell({r:headerRow,c:range.s.c})];const routeValue=String(routeCell?.w??routeCell?.v??'').trim();const routeMerge=anchors.get(`${headerRow},${range.s.c}`);if(routeMerge&&routeMerge.e.c>=range.s.c+1&&formatDateText(routeValue)!==routeValue){currentGroup=formatDateText(routeValue);const styleId=cellStyleIds.get(sheet.name)?.get(XLSX.utils.encode_cell({r:headerRow,c:range.s.c}))??0;currentGroupStyle=groupHeadingStyle(styleDefinitions[styleId]||{});activeGroup=currentGroup}}
  for(let r=headerRow+1;r<=range.e.r;r++){
    if(ws['!rows']?.[r]?.hidden)continue;
    if(isRoute){const routeCell=ws[XLSX.utils.encode_cell({r,c:range.s.c})];const routeValue=String(routeCell?.w??routeCell?.v??'').trim();const routeMerge=anchors.get(`${r},${range.s.c}`);if(routeMerge&&routeMerge.e.c>=range.s.c+1&&formatDateText(routeValue)!==routeValue){currentGroup=formatDateText(routeValue);const routeStyleId=cellStyleIds.get(sheet.name)?.get(XLSX.utils.encode_cell({r,c:range.s.c}))??0;currentGroupStyle=groupHeadingStyle(styleDefinitions[routeStyleId]||{});continue}}
    if(groupColumn!==null){const rawGroup=effectiveCellValue(ws,r,groupColumn,anchors,covered,headerRow);if(rawGroup){const normalized=groupColumn===0&&sheet.name!=='ΠΑΡΑΓΓΕΛΙΕΣ'?formatDateText(rawGroup):rawGroup;if(normalized!==activeGroup){activeGroup=normalized;currentGroup=normalized;currentGroupStyle=''}}}
    const fields=[];let rowText='';let recordTitle='';const seen=new Set();
    for(let c=range.s.c;c<=range.e.c;c++){
      let sourceRow=r,sourceCol=c;const keyHere=`${r},${c}`;
      if(covered.has(keyHere)){
        const merge=mergesForCell(anchors,r,c);if(!merge)continue;
        sourceRow=merge.s.r;sourceCol=merge.s.c;
        if(sourceRow!==r||seen.has(`${sourceRow},${sourceCol}`))continue;
        seen.add(`${sourceRow},${sourceCol}`);
      }
      const address=XLSX.utils.encode_cell({r:sourceRow,c:sourceCol});const cell=ws[address]||{};
      const rawValue=dirty.get(`${sheet.name}!${address}`)?.value??cell.w??cell.v??'';const value=formatCellValue(cell,rawValue,headerForColumn(ws,headerRow,c));if(!String(value).trim())continue;
      const headerCell=ws[XLSX.utils.encode_cell({r:headerRow,c})];
      const header=String(headerCell?.w??headerCell?.v??'').trim()||XLSX.utils.encode_col(c);
      if(!recordTitle)recordTitle=value;
      rowText+=` ${header} ${value}`;
      const styleId=cellStyleIds.get(sheet.name)?.get(address)??0;const style=styleDefinitions[styleId]||styleDefinitions[0]||{};
      const locked=sheet.formulaCells.has(address);const editable=editMode&&!locked&&sourceRow===r;const changed=dirty.has(`${sheet.name}!${address}`);
      const content=cell.l?.Target?hyperlinkMarkup(cell.l,value):esc(value);
      const hit=query&&rowText.toLocaleLowerCase('el').includes(query)?' search-match':'';
      fields.push(`<div class="mobile-field"><span class="mobile-label">${esc(header)}</span><div class="mobile-value edit-cell ${locked?'locked ':''}${changed?'changed':''}${hit}" ${editable?'contenteditable="true" spellcheck="false"':''} data-address="${address}" data-sheet="${esc(sheet.name)}" data-header="${esc(header)}" title="${locked?'Τύπος Excel — δεν επεξεργάζεται από την εφαρμογή':cell.l?.Tooltip|| (editable?'Πάτησε για επεξεργασία':'')}" style="${styleCss(style)}">${content}</div></div>`);
    }
    if(fields.length&&(!query||rowText.toLocaleLowerCase('el').includes(query))){const group=currentGroup?`<div class="mobile-group" style="${currentGroupStyle}">${esc(currentGroup)}</div>`:'';records.push(`${group}<article class="mobile-record"><div class="mobile-record-title">${esc(recordTitle||`Γραμμή ${r+1}`)}</div>${fields.join('')}</article>`);if(group)currentGroup=''}
  }
  area.className='mobile-records'+(editMode?' editing':'');
  area.innerHTML=records.join('')||'<div class="empty">Δεν βρέθηκαν γραμμές με περιεχόμενο.</div>';
  bindEditedCells();
  area.querySelectorAll('a').forEach(link=>link.addEventListener('click',event=>{if(editMode)event.preventDefault()}));
  if(scrollToEnd)requestAnimationFrame(()=>area.lastElementChild?.scrollIntoView({block:'end'}));
}
function headerForColumn(ws,row,col){const cell=ws[XLSX.utils.encode_cell({r:row,c:col})];return String(cell?.w??cell?.v??'')}
function groupHeadingStyle(style){const css=[];if(style.font?.color)css.push(`color:${style.font.color}`);else css.push('color:#c62828');css.push('font-weight:800');if(style.font?.size)css.push(`font-size:${Math.max(10,Math.min(22,style.font.size))}pt`);return css.join(';')}
function effectiveCellValue(ws,row,col,anchors,covered,headerRow){let r=row,c=col;if(covered.has(`${row},${col}`)){const merge=mergesForCell(anchors,row,col);if(merge){r=merge.s.r;c=merge.s.c}}const cell=ws[XLSX.utils.encode_cell({r,c})];return String(formatCellValue(cell,cell?.v??cell?.w??'',headerForColumn(ws,headerRow,col))).trim()}
function formatCellValue(cell,raw,header=''){if(raw==null||raw==='')return '';const isDate=XLSX.SSF?.is_date?.(cell?.z)||/ΗΜΕΡΟΜΗΝΙΑ|ΠΑΡΑΔΟΣΗ|ΟΛΟΚΛΗΡΩΣΗ/i.test(header);if(isDate&&typeof cell?.v==='number'&&(typeof raw==='number'||raw===cell.w||raw===cell.v)){const date=XLSX.SSF.parse_date_code(cell.v);if(date?.y&&date?.m&&date?.d)return`${String(date.d).padStart(2,'0')}/${String(date.m).padStart(2,'0')}/${date.y}`;}return formatDateText(String(raw))}
function formatDateText(value){return String(value).replace(/(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})(?!\d)/g,(_,day,month,year)=>`${day.padStart(2,'0')}/${month.padStart(2,'0')}/${year.length===2?`20${year}`:year}`)}
function mergesForCell(anchors,r,c){for(const merge of anchors.values())if(r>=merge.s.r&&r<=merge.e.r&&c>=merge.s.c&&c<=merge.e.c)return merge;return null}
function bindEditedCells(){area.querySelectorAll('[contenteditable="true"]').forEach(cell=>cell.addEventListener('input',()=>{const key=`${cell.dataset.sheet}!${cell.dataset.address}`;dirty.set(key,{sheet:cell.dataset.sheet,address:cell.dataset.address,value:cell.textContent,header:cell.dataset.header});cell.classList.add('changed');saveButton.disabled=false}))}
function renderControls(){editToggle.textContent=editMode?'✎ Τέλος επεξεργασίας':'✎ Επεξεργασία';saveButton.hidden=!editMode;cancelButton.hidden=!editMode;saveButton.disabled=dirty.size===0;editToggle.hidden=!remoteFile}

async function saveChanges(){
  if(!remoteFile||dirty.size===0)return;
  saveButton.disabled=true;setNotice('Έλεγχος για νεότερη έκδοση και αποθήκευση…');
  try{
    const token=await getAccessToken(localStorage.getItem(CLIENT_KEY));
    const zip=await JSZip.loadAsync(originalBytes);
    const sheetPathMap=await getSheetPaths(zip);
    const changesBySheet=new Map();for(const change of dirty.values()){if(!changesBySheet.has(change.sheet))changesBySheet.set(change.sheet,[]);changesBySheet.get(change.sheet).push(change)}
    for(const [sheetName,changes] of changesBySheet){const path=sheetPathMap.get(sheetName);if(!path)throw new Error(`Δεν βρέθηκε το XML του φύλλου ${sheetName}.`);const xml=await zip.file(path).async('string');zip.file(path,patchSheetXml(xml,changes))}
    const updatedBytes=await zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
    const current=await graphJson(`${GRAPH}/drives/${encodeURIComponent(remoteFile.driveId)}/items/${encodeURIComponent(remoteFile.id)}?$select=id,eTag`,token);
    if(remoteFile.eTag&&current.eTag&&remoteFile.eTag!==current.eTag){setNotice('Το αρχείο άλλαξε στο OneDrive μετά τη φόρτωσή του. Δεν αποθήκευσα τις αλλαγές σου για να μην αντικαταστήσω νεότερα δεδομένα. Κάνε ξανά σύνδεση και επανάλαβε τις αλλαγές.');saveButton.disabled=false;return}
    const response=await graphFetch(`${GRAPH}/drives/${encodeURIComponent(remoteFile.driveId)}/items/${encodeURIComponent(remoteFile.id)}/content`,token,{method:'PUT',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},body:updatedBytes});
    if(!response.ok)throw new Error(`Η αποθήκευση απέτυχε (${response.status}). Βεβαιώσου ότι έχεις δικαίωμα επεξεργασίας στο κοινόχρηστο αρχείο.`);
    const saved=await response.json();remoteFile.eTag=saved.eTag||current.eTag;originalBytes=updatedBytes;dirty.clear();editMode=false;
    await loadBytes(originalBytes,remoteFile.name);setNotice('Οι αλλαγές αποθηκεύτηκαν στο κοινόχρηστο Excel.');
  }catch(error){saveButton.disabled=false;setNotice(error.message||'Δεν έγινε αποθήκευση.');console.error(error)}
}

async function getSheetPaths(zip){
  const workbook=xmlDoc(await zip.file('xl/workbook.xml').async('string'));
  const relationships=xmlDoc(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
  const rels=new Map([...relationships.getElementsByTagNameNS(PKGNS,'Relationship')].map(rel=>[rel.getAttribute('Id'),rel.getAttribute('Target')]));
  const paths=new Map();
  for(const sheet of workbook.getElementsByTagNameNS(NS,'sheet')){const target=rels.get(sheet.getAttributeNS(RELNS,'id'));if(!target)continue;const path=target.startsWith('/')?target.slice(1):target.startsWith('xl/')?target:`xl/${target}`;paths.set(sheet.getAttribute('name'),path)}
  return paths;
}
function patchSheetXml(source,changes){
  const doc=xmlDoc(source);const data=doc.getElementsByTagNameNS(NS,'sheetData')[0];if(!data)throw new Error('Το φύλλο δεν έχει έγκυρα δεδομένα Excel.');
  for(const change of changes){
    const rowNumber=Number(change.address.match(/\d+$/)?.[0]);if(!rowNumber)throw new Error('Μη έγκυρη διεύθυνση κελιού.');
    let row=[...data.getElementsByTagNameNS(NS,'row')].find(node=>Number(node.getAttribute('r'))===rowNumber);
    if(!row){row=doc.createElementNS(NS,'row');row.setAttribute('r',String(rowNumber));const next=[...data.getElementsByTagNameNS(NS,'row')].find(node=>Number(node.getAttribute('r'))>rowNumber);data.insertBefore(row,next||null)}
    let cell=[...row.getElementsByTagNameNS(NS,'c')].find(node=>node.getAttribute('r')===change.address);
    if(!cell){cell=doc.createElementNS(NS,'c');cell.setAttribute('r',change.address);const col=XLSX.utils.decode_cell(change.address).c;const next=[...row.getElementsByTagNameNS(NS,'c')].find(node=>XLSX.utils.decode_cell(node.getAttribute('r')).c>col);row.insertBefore(cell,next||null)}
    if(cell.getElementsByTagNameNS(NS,'f').length)throw new Error(`Το κελί ${change.address} περιέχει τύπο Excel και δεν αλλάζει από την εφαρμογή.`);
    setCellValue(doc,cell,change.value,change.header);
  }
  return new XMLSerializer().serializeToString(doc);
}
function setCellValue(doc,cell,raw,header){
  for(const child of [...cell.childNodes])if(child.nodeType===1&&(child.localName==='v'||child.localName==='is'||child.localName==='f'))cell.removeChild(child);
  const value=String(raw??'').trim();
  if(!value){cell.removeAttribute('t');return}
  if(/ΗΜΕΡΟΜΗΝΙΑ|ΠΑΡΑΔΟΣΗ|ΟΛΟΚΛΗΡΩΣΗ/i.test(header)){const serial=toExcelDate(value);if(serial!==null){cell.removeAttribute('t');const v=doc.createElementNS(NS,'v');v.textContent=String(serial);cell.appendChild(v);return}}
  const number=Number(value.replace(',','.'));
  if(value!==''&&Number.isFinite(number)&&/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(value.replace(',','.'))){cell.removeAttribute('t');const v=doc.createElementNS(NS,'v');v.textContent=String(number);cell.appendChild(v);return}
  cell.setAttribute('t','inlineStr');const inline=doc.createElementNS(NS,'is');const text=doc.createElementNS(NS,'t');text.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');text.textContent=value;inline.appendChild(text);cell.appendChild(inline);
}
function toExcelDate(value){const match=value.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);if(!match)return null;const day=Number(match[1]),month=Number(match[2]),year=Number(match[3]);const date=new Date(Date.UTC(year,month-1,day));if(date.getUTCDate()!==day||date.getUTCMonth()!==month-1)return null;return Math.round((date-Date.UTC(1899,11,30))/86400000)}
function xmlDoc(source){const doc=new DOMParser().parseFromString(source,'application/xml');if(doc.getElementsByTagName('parsererror').length)throw new Error('Δεν ήταν δυνατή η ανάγνωση της δομής του Excel.');return doc}

function setNotice(message){notice.textContent=message}
function fail(error){console.error(error);setNotice(error?.message||'Παρουσιάστηκε σφάλμα.');alert(error?.message||'Παρουσιάστηκε σφάλμα.')}
function esc(value){return String(value).replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]))}
if('serviceWorker'in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('./sw.js').catch(()=>{});
