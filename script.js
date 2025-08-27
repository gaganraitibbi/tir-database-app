/* ========= Final cleaned script.js (UPDATED: Encumbrance + Advocate in Preview/Excel) =========
   Base file used: user's uploaded script.js. Modifications:
   - Added buildEncumbranceText(st)
   - renderPreview() uses buildEncumbranceText and st.advocates
   - _buildExcelRowFromState uses buildEncumbranceText(st) for Encumbrance column
   (Other original logic preserved)
*/

/* ========= Helpers ========= */
function num(x){ const n = parseFloat(x); return isNaN(n)?0:n; }
function $(selector, ctx=document){ try{ return ctx.querySelector(selector); } catch(e){ return null; } }
function $$(selector, ctx=document){ try{ return Array.from(ctx.querySelectorAll(selector)); } catch(e){ return []; } }

const LS_FORM = "TIR_FORM_MULTI";
const LS_ENTRIES = "TIR_ENTRIES";

/* parseShareFactor: accepts "1/2", "50%", "0.5", "50" -> returns number between 0 and 1 */
function parseShareFactor(str){
  if(str === undefined || str === null) return 1;
  const s = String(str).trim();
  if(s === "") return 1;

  var frac = s.match(/^(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)$/);
  if(frac){
    var a = parseFloat(frac[1]), b = parseFloat(frac[2]);
    if(!isNaN(a) && !isNaN(b) && b !== 0) return Math.max(0, Math.min(1, a/b));
    return 1;
  }
  var perc = s.match(/^(\d+(?:\.\d+)?)\s*%$/);
  if(perc){
    var p = parseFloat(perc[1]);
    if(!isNaN(p)) return Math.max(0, Math.min(1, p/100));
    return 1;
  }
  var n = parseFloat(s);
  if(isNaN(n)) return 1;
  if(n > 1.000001) return Math.max(0, Math.min(1, n/100));
  return Math.max(0, Math.min(1, n));
}

/* Convert hectares to "X.XXX Acres, Y Bigha & Z Biswa" */
function hectaresToAcresText(hect){
  var h = parseFloat(hect) || 0;
  var acres = h * 2.47105381;
  var totalBigha = acres / 0.625;
  var bigha = Math.floor(totalBigha);
  var biswa = Math.round((totalBigha - bigha) * 20);
  if(biswa === 20){ bigha += 1; biswa = 0; }
  return { acres: acres, text: acres.toFixed(3) + " Acres, " + bigha + " Bigha & " + biswa + " Biswa" };
}

/* ========= Property Set UI (Templates assumed in HTML) ========= */
function addPropertySet(data){
  var wrap = $("#propertySetContainer");
  var tpl  = $("#tplPropertySet");
  if(!wrap || !tpl){ console.warn("propertySetContainer or tplPropertySet missing"); return null; }
  var setNode = tpl.content.firstElementChild.cloneNode(true);

  setNode.addEventListener("input", function(e){
    var trg = e.target;
    if(!trg) return;
    if(trg.name === "area" || trg.name && trg.name.indexOf("prop") === 0 || trg.classList && trg.classList.contains("borrowerShare")){
      recalcSetTotals(setNode);
      recalcGrandTotals();
      saveState();
    }
  });

  if(data){
    var el;
    el = $(".fixedChak", setNode); if(el) el.value = data.fixed?.chak || "";
    el = $(".fixedTehsil", setNode); if(el) el.value = data.fixed?.tehsil || "";
    el = $(".fixedDistt", setNode); if(el) el.value = data.fixed?.distt || "";
    el = $(".fixedSamvat", setNode); if(el) el.value = data.fixed?.samvat || "";
    el = $(".fixedAcno", setNode); if(el) el.value = data.fixed?.acno || "";

    el = $("input[name='propTotal']", setNode); if(el) el.value = data.totals?.total || "";
    el = $("input[name='propCommand']", setNode); if(el) el.value = data.totals?.command || "";
    el = $("input[name='propUncommand']", setNode); if(el) el.value = data.totals?.uncommand || "";
    el = $("input[name='propKhala']", setNode); if(el) el.value = data.totals?.khala || "";
    el = $("input[name='propRasta']", setNode); if(el) el.value = data.totals?.rasta || "";
    el = $("input[name='propKuan']", setNode); if(el) el.value = data.totals?.kuan || "";
    el = $("input[name='propNahari']", setNode); if(el) el.value = data.totals?.nahari || "";
    el = $("input[name='propNali1st']", setNode); if(el) el.value = data.totals?.nali1 || "";
    el = $("input[name='propNali2nd']", setNode); if(el) el.value = data.totals?.nali2 || "";
    el = $("input[name='propBarani1st']", setNode); if(el) el.value = data.totals?.barani1 || "";

    // rows
    var group = $(".sno-group", setNode);
    if(group){ group.innerHTML = ""; (data.rows || []).forEach(function(r){ addSNoGroup(setNode, r); }); }
    var b = $(".borrowerShare", setNode); if(b) b.value = data.borrowerShare || "";
  }

  wrap.appendChild(setNode);
  recalcSetTotals(setNode);
  recalcGrandTotals();
  return setNode;
}

/* ===== rows helpers ===== */
function addSNoGroup(setNode, rowData){
  var tpl = $("#tplSNoGroup");
  if(!tpl){ console.warn("tplSNoGroup missing"); return null; }
  var row = tpl.content.firstElementChild.cloneNode(true);

  var btn;
  btn = $(".btnAddKNo", row); if(btn) btn.addEventListener("click", function(){ var newRow = addKNoRow(setNode, null, row); saveState(); var inp = $("input[name='kno']", newRow); if(inp) inp.focus(); });
  btn = $(".btnAddSNo", row); if(btn) btn.addEventListener("click", function(){ addSNoGroup(setNode); saveState(); });
  btn = $(".btnRemoveRow", row); if(btn) btn.addEventListener("click", function(){ row.remove(); recalcSetTotals(setNode); recalcGrandTotals(); saveState(); });
  var areaInp = $("input[name='area']", row); if(areaInp) areaInp.addEventListener("input", function(){ recalcSetTotals(setNode); recalcGrandTotals(); });

  if(rowData){
    var e;
    e = $("input[name='sno']", row); if(e) e.value = rowData.sno || "";
    e = $("input[name='murba']", row); if(e) e.value = rowData.murba || "";
    e = $("input[name='kno']", row); if(e) e.value = rowData.kno || "";
    e = $("input[name='area']", row); if(e) e.value = rowData.area || "";
  }

  var g = $(".sno-group", setNode); if(g) g.appendChild(row);
  return row;
}

function addKNoRow(setNode, data, anchorRow){
  var tpl = $("#tplKNoOnly");
  if(!tpl){ console.warn("tplKNoOnly missing"); return null; }
  var row = tpl.content.firstElementChild.cloneNode(true);

  var btn;
  btn = $(".btnAddKNo", row); if(btn) btn.addEventListener("click", function(){ var newRow = addKNoRow(setNode, null, row); saveState(); var inp = $("input[name='kno']", newRow); if(inp) inp.focus(); });
  btn = $(".btnAddSNo", row); if(btn) btn.addEventListener("click", function(){ addSNoGroup(setNode); saveState(); });
  btn = $(".btnRemoveRow", row); if(btn) btn.addEventListener("click", function(){ row.remove(); recalcSetTotals(setNode); recalcGrandTotals(); saveState(); });
  var areaInp = $("input[name='area']", row); if(areaInp) areaInp.addEventListener("input", function(){ recalcSetTotals(setNode); recalcGrandTotals(); });

  if(data){
    var e = $("input[name='kno']", row); if(e) e.value = data.kno || "";
    e = $("input[name='area']", row); if(e) e.value = data.area || "";
  }

  if(anchorRow && anchorRow.parentNode) anchorRow.insertAdjacentElement("afterend", row);
  else { var g = $(".sno-group", setNode); if(g) g.appendChild(row); }
  return row;
}

/* ========= Per-set totals ========= */
function recalcSetTotals(setNode){
  if(!setNode) return;
  var areaTotal = 0;
  $$("input[name='area']", setNode).forEach(function(inp){ areaTotal += parseFloat(inp.value || "0") || 0; });
  var totalInput = $("input[name='propTotal']", setNode); if(totalInput) totalInput.value = areaTotal.toFixed(4);

  var cmd   = parseFloat($("input[name='propCommand']", setNode)?.value || "0") || 0;
  var uncmd = parseFloat($("input[name='propUncommand']", setNode)?.value || "0") || 0;
  var khala = parseFloat($("input[name='propKhala']", setNode)?.value || "0") || 0;
  var rasta = parseFloat($("input[name='propRasta']", setNode)?.value || "0") || 0;
  var kuan  = parseFloat($("input[name='propKuan']", setNode)?.value || "0") || 0;
  var nahari= parseFloat($("input[name='propNahari']", setNode)?.value || "0") || 0;
  var nali1 = parseFloat($("input[name='propNali1st']", setNode)?.value || "0") || 0;
  var nali2 = parseFloat($("input[name='propNali2nd']", setNode)?.value || "0") || 0;
  var bar1  = parseFloat($("input[name='propBarani1st']", setNode)?.value || "0") || 0;

  var factor = parseShareFactor(($(".borrowerShare", setNode) ? $(".borrowerShare", setNode).value : ""));
  var bTotal   = areaTotal * factor;
  var bCommand = cmd       * factor;
  var bUncmd   = uncmd     * factor;
  var bKhala   = khala     * factor;
  var bRasta   = rasta     * factor;
  var bKuan    = kuan      * factor;
  var bNahari  = nahari    * factor;
  var bNali1   = nali1     * factor;
  var bNali2   = nali2     * factor;
  var bBar1    = bar1      * factor;

  var el;
  el = $(".tblTotalLand", setNode); if(el) el.textContent = bTotal.toFixed(4);
  el = $(".tblCommand", setNode); if(el) el.textContent = bCommand.toFixed(4);
  el = $(".tblUncommand", setNode); if(el) el.textContent = bUncmd.toFixed(4);
  el = $(".tblKhala", setNode); if(el) el.textContent = bKhala.toFixed(4);
  el = $(".tblRasta", setNode); if(el) el.textContent = bRasta.toFixed(4);
  el = $(".tblKuan", setNode); if(el) el.textContent = bKuan.toFixed(4);
  el = $(".tblNahari", setNode); if(el) el.textContent = bNahari.toFixed(4);
  el = $(".tblNali1", setNode); if(el) el.textContent = bNali1.toFixed(4);
  el = $(".tblNali2", setNode); if(el) el.textContent = bNali2.toFixed(4);
  el = $(".tblBarani1", setNode); if(el) el.textContent = bBar1.toFixed(4);

  var bAcre = bTotal * 2.47105381;
  el = $(".tblAcre", setNode); if(el) el.textContent = bAcre.toFixed(3);
  var bbBigha = Math.floor(bAcre/0.625);
  var bbBiswa = Math.floor(((bAcre/0.625) - bbBigha) * 20);
  el = $(".tblBighaBiswa", setNode); if(el) el.textContent = bbBigha + "/" + bbBiswa;
}

/* ========= Grand totals ========= */
function recalcGrandTotals(){
  function T(id, val){ var e = document.getElementById(id); if(e) e.textContent = val; }

  var total=0, command=0, uncommand=0, khala=0, rasta=0, kuan=0, nahari=0, nali1=0, nali2=0, bar1=0;
  var shareTotal=0, shareCommand=0, shareUncommand=0, shareKhala=0, shareRasta=0, shareKuan=0, shareNahari=0, shareNali1=0, shareNali2=0, shareBar1=0;

  $$("#propertySetContainer .property-set").forEach(function(set){
    var areaTotal = parseFloat($("input[name='propTotal']", set)?.value || "0") || 0;
    var cmd = parseFloat($("input[name='propCommand']", set)?.value || "0") || 0;
    var uncmd = parseFloat($("input[name='propUncommand']", set)?.value || "0") || 0;
    var kh = parseFloat($("input[name='propKhala']", set)?.value || "0") || 0;
    var rs = parseFloat($("input[name='propRasta']", set)?.value || "0") || 0;
    var ku = parseFloat($("input[name='propKuan']", set)?.value || "0") || 0;
    var nh = parseFloat($("input[name='propNahari']", set)?.value || "0") || 0;
    var n1 = parseFloat($("input[name='propNali1st']", set)?.value || "0") || 0;
    var n2 = parseFloat($("input[name='propNali2nd']", set)?.value || "0") || 0;
    var b1 = parseFloat($("input[name='propBarani1st']", set)?.value || "0") || 0;

    total += areaTotal;
    command += cmd;
    uncommand += uncmd;
    khala += kh;
    rasta += rs;
    kuan += ku;
    nahari += nh;
    nali1 += n1;
    nali2 += n2;
    bar1 += b1;

    var factor = parseShareFactor(($(".borrowerShare", set) ? $(".borrowerShare", set).value : ""));
    if(factor > 0){
      shareTotal += areaTotal * factor;
      shareCommand += cmd * factor;
      shareUncommand += uncmd * factor;
      shareKhala += kh * factor;
      shareRasta += rs * factor;
      shareKuan += ku * factor;
      shareNahari += nh * factor;
      shareNali1 += n1 * factor;
      shareNali2 += n2 * factor;
      shareBar1 += b1 * factor;
    }
  });

  T("gTotalLand", total.toFixed(4));
  T("gCommand", command.toFixed(4));
  T("gUncommand", uncommand.toFixed(4));
  T("gKhala", khala.toFixed(4));
  T("gRasta", rasta.toFixed(4));
  T("gKuan", kuan.toFixed(4));
  T("gNahari", nahari.toFixed(4));
  T("gNali1", nali1.toFixed(4));
  T("gNali2", nali2.toFixed(4));
  T("gBarani1", bar1.toFixed(4));

  T("gShareTotal", shareTotal.toFixed(4));
  T("gShareCommand", shareCommand.toFixed(4));
  T("gShareUncommand", shareUncommand.toFixed(4));
  T("gShareKhala", shareKhala.toFixed(4));
  T("gShareRasta", shareRasta.toFixed(4));
  T("gShareKuan", shareKuan.toFixed(4));
  T("gShareNahari", shareNahari.toFixed(4));
  T("gShareNali1", shareNali1.toFixed(4));
  T("gShareNali2", shareNali2.toFixed(4));
  T("gShareBarani1", shareBar1.toFixed(4));

  var gAcre = total * 2.47105381;
  T("gAcre", gAcre.toFixed(3));
  var gBigha = Math.floor(gAcre/0.625);
  var gBiswa = Math.floor(((gAcre/0.625)-gBigha)*20);
  T("gBigha", String(gBigha));
  T("gBiswa", String(gBiswa));

  var sAcre = shareTotal * 2.47105381;
  T("gShareAcre", sAcre.toFixed(3));
  var sBigha = Math.floor(sAcre/0.625);
  var sBiswa = Math.floor(((sAcre/0.625)-sBigha)*20);
  T("gShareBigha", String(sBigha));
  T("gShareBiswa", String(sBiswa));
}

/* ========= Navigation ========= */
function goTo(step){
  $$(".step").forEach(function(s){ s.classList.remove("active"); });
  var dest = document.getElementById("step"+step);
  if(dest) dest.classList.add("active");
  $$(".stepper .step-node").forEach(function(b){ b.classList.remove("active"); });
  var node = document.querySelector('.stepper .step-node[data-goto="'+step+'"]');
  if(node) node.classList.add("active");
}

/* ========= State ========= */
function collectState(){
  var st = {};
  st.srNo = $("#srNo") ? $("#srNo").value : "";
  st.bankDate = $("#bankDate") ? $("#bankDate").value : "";
  st.bankName = $("#bankName") ? $("#bankName").value : "";
  st.branchHindi = $("#branchHindi") ? $("#branchHindi").value : "";
  st.branchEnglish = $("#branchEnglish") ? $("#branchEnglish").value : "";
  st.nameHindi = $("#nameHindi") ? $("#nameHindi").value : "";
  st.nameEnglish = $("#nameEnglish") ? $("#nameEnglish").value : "";

  st.propertySets = [];
  $$("#propertySetContainer .property-set").forEach(function(set){
    var rows = $$(".property-row", set).map(function(r){ return {
      sno: $("input[name='sno']", r) ? $("input[name='sno']", r).value : "",
      murba: $("input[name='murba']", r) ? $("input[name='murba']", r).value : "",
      kno: $("input[name='kno']", r) ? $("input[name='kno']", r).value : "",
      area: $("input[name='area']", r) ? $("input[name='area']", r).value : ""
    }; });
    var totals = {
      total: $("input[name='propTotal']", set) ? $("input[name='propTotal']", set).value : "",
      command: $("input[name='propCommand']", set) ? $("input[name='propCommand']", set).value : "",
      uncommand: $("input[name='propUncommand']", set) ? $("input[name='propUncommand']", set).value : "",
      khala: $("input[name='propKhala']", set) ? $("input[name='propKhala']", set).value : "",
      rasta: $("input[name='propRasta']", set) ? $("input[name='propRasta']", set).value : "",
      kuan: $("input[name='propKuan']", set) ? $("input[name='propKuan']", set).value : "",
      nahari: $("input[name='propNahari']", set) ? $("input[name='propNahari']", set).value : "",
      nali1: $("input[name='propNali1st']", set) ? $("input[name='propNali1st']", set).value : "",
      nali2: $("input[name='propNali2nd']", set) ? $("input[name='propNali2nd']", set).value : "",
      barani1: $("input[name='propBarani1st']", set) ? $("input[name='propBarani1st']", set).value : ""
    };
    st.propertySets.push({
      fixed: {
        chak: $(".fixedChak", set) ? $(".fixedChak", set).value : "",
        tehsil: $(".fixedTehsil", set) ? $(".fixedTehsil", set).value : "",
        distt: $(".fixedDistt", set) ? $(".fixedDistt", set).value : "",
        samvat: $(".fixedSamvat", set) ? $(".fixedSamvat", set).value : "",
        acno: $(".fixedAcno", set) ? $(".fixedAcno", set).value : ""
      },
      borrowerShare: $(".borrowerShare", set) ? $(".borrowerShare", set).value : "",
      rows: rows,
      totals: totals
    });
  });

  st.totals = {
    total: document.getElementById("gTotalLand") ? document.getElementById("gTotalLand").textContent : "0.0000",
    command: document.getElementById("gCommand") ? document.getElementById("gCommand").textContent : "0.0000",
    uncommand: document.getElementById("gUncommand") ? document.getElementById("gUncommand").textContent : "0.0000",
    acre: document.getElementById("gAcre") ? document.getElementById("gAcre").textContent : "0.000",
    bigha: document.getElementById("gBigha") ? document.getElementById("gBigha").textContent : "0",
    biswa: document.getElementById("gBiswa") ? document.getElementById("gBiswa").textContent : "0"
  };

  st.boundaries = {
    north: $("#boundaryNorth") ? $("#boundaryNorth").value : "",
    south: $("#boundarySouth") ? $("#boundarySouth").value : "",
    east: $("#boundaryEast") ? $("#boundaryEast").value : "",
    west: $("#boundaryWest") ? $("#boundaryWest").value : ""
  };

  st.encumbrance = {
    status: $("#encumbranceStatus") ? $("#encumbranceStatus").value : "",
    manual: $("#encumbranceManual") ? $("#encumbranceManual").value : ""
  };

  st.srDetails = [];
  $$("#srDetailContainer .sr-detail-row").forEach(function(r){
    st.srDetails.push({
      office: $(".srOffice", r) ? $(".srOffice", r).value : "",
      receipt: $(".srReceipt", r) ? $(".srReceipt", r).value : "",
      grn: $(".srGrn", r) ? $(".srGrn", r).value : ""
    });
  });

  st.searchDetail = {
    yearFrom: $("#srYearFrom") ? $("#srYearFrom").value : "",
    yearTo: $("#srYearTo") ? $("#srYearTo").value : "",
    amount: $("#searchAmount") ? $("#searchAmount").value : "",
    giftOther: $("#giftOther") ? $("#giftOther").value : "",
    dlcRate: $("#dlcRate") ? $("#dlcRate").value : ""
  };

  st.documents = {
    jamabandi: $("#docJamabandi") ? $("#docJamabandi").value : "",
    girdawari: $("#docGirdawari") ? $("#docGirdawari").value : "",
    map: $("#docMap") ? $("#docMap").value : "",
    samvat: $("#docSamvat") ? $("#docSamvat").value : "",
    other: $("#docOther") ? $("#docOther").value : ""
  };

  st.advocates = [];
  $$("#advocateContainer .advocate-row").forEach(function(r){
    st.advocates.push({
      name: $("#advocateName", r) ? $("#advocateName", r).value : "",
      mobile: $("#advocateMobile", r) ? $("#advocateMobile", r).value : "",
      chamber: $("#advocateChamber", r) ? $("#advocateChamber", r).value : "",
      other: $("#advocateOtherInput", r) ? $("#advocateOtherInput", r).value : ""
    });
  });

  return st;
}

function saveState(){ try{ localStorage.setItem(LS_FORM, JSON.stringify(collectState())); }catch(e){ console.warn("saveState failed", e); } }

function loadState(){
  var s = localStorage.getItem(LS_FORM);
  var container = $("#propertySetContainer");
  if(!container){ console.warn("propertySetContainer not found"); return; }
  container.innerHTML = "";
  if(!s){ addPropertySet(); recalcGrandTotals(); return; }
  var st;
  try{ st = JSON.parse(s); }catch(e){ console.warn("invalid draft json", e); addPropertySet(); recalcGrandTotals(); return; }

  var el;
  el = $("#srNo"); if(el) el.value = st.srNo || "";
  el = $("#bankDate"); if(el) el.value = st.bankDate || "";
  el = $("#bankName"); if(el) el.value = st.bankName || "";
  el = $("#branchHindi"); if(el) el.value = st.branchHindi || "";
  el = $("#branchEnglish"); if(el) el.value = st.branchEnglish || "";
  el = $("#nameHindi"); if(el) el.value = st.nameHindi || "";
  el = $("#nameEnglish"); if(el) el.value = st.nameEnglish || "";

  if(st.propertySets && st.propertySets.length){
    st.propertySets.forEach(function(ps){ addPropertySet(ps); });
  } else { addPropertySet(); }

  if(st.boundaries){
    el = $("#boundaryNorth"); if(el) el.value = st.boundaries.north || "";
    el = $("#boundarySouth"); if(el) el.value = st.boundaries.south || "";
    el = $("#boundaryEast"); if(el) el.value = st.boundaries.east || "";
    el = $("#boundaryWest"); if(el) el.value = st.boundaries.west || "";
  }

  if(st.encumbrance){
    el = $("#encumbranceStatus"); if(el) el.value = st.encumbrance.status || "";
    el = $("#encumbranceManual"); if(el) el.value = st.encumbrance.manual || "";
    toggleEncumbranceManual();
  }

  $("#srDetailContainer").innerHTML = "";
  if(st.srDetails && st.srDetails.length){
    st.srDetails.forEach(function(d){
      var row = addSRDetailRow();
      var e = $(".srOffice", row); if(e) e.value = d.office || "";
      e = $(".srReceipt", row); if(e) e.value = d.receipt || "";
      e = $(".srGrn", row); if(e) e.value = d.grn || "";
    });
  } else { addSRDetailRow(); }

  if(st.searchDetail){
    el = $("#srYearFrom"); if(el) el.value = st.searchDetail.yearFrom || "";
    el = $("#srYearTo"); if(el) el.value = st.searchDetail.yearTo || "";
    el = $("#searchAmount"); if(el) el.value = st.searchDetail.amount || "";
    el = $("#giftOther"); if(el) el.value = st.searchDetail.giftOther || "";
    el = $("#dlcRate"); if(el) el.value = st.searchDetail.dlcRate || "";
  }

  if(st.documents){
    el = $("#docJamabandi"); if(el) el.value = st.documents.jamabandi || "";
    el = $("#docGirdawari"); if(el) el.value = st.documents.girdawari || "";
    el = $("#docMap"); if(el) el.value = st.documents.map || "";
    el = $("#docSamvat"); if(el) el.value = st.documents.samvat || "";
    el = $("#docOther"); if(el) el.value = st.documents.other || "";
  }

  // restore advocates into DOM rows if needed (optional)
  // Many templates create an advocate row on button click; we don't auto-create here to avoid UI duplication
  recalcGrandTotals();
}

/* ========= Encumbrance helper ========= */
function toggleEncumbranceManual(){
  var wrap = $("#encumbranceManualWrap");
  if(!wrap) return;
  var st = $("#encumbranceStatus"); if(st && (st.value === "mortgagedPartial" || st.value === "3")) wrap.style.display = "block"; else wrap.style.display = "none";
}

/* NEW: buildEncumbranceText - use in Preview + Excel so both match */
function buildEncumbranceText(st){
  var encText = "-";
  if(!st || !st.encumbrance || !st.encumbrance.status) return encText;
  var bank = st.bankName || "_____";
  var branch = st.branchEnglish || st.branchHindi || "_____";
  var status = st.encumbrance.status;
  if(status === "mortgagedFull" || status === "mortgaged" || status === "1"){
    encText = "The Land/Share of Borrower (s) is already mortgaged in favour of " + bank + " Branch " + branch + ".";
  } else if(status === "free" || status === "2"){
    encText = "The Land/Share of Borrower (s) is free from all encumbrances.";
  } else if(status === "mortgagedPartial" || status === "3"){
    var manual = st.encumbrance.manual || "{without Land of Chak/Stone/Killa../Area}";
    encText = "The Land / Share of Borrower " + manual + " is already mortgaged in favour of " + bank + " Branch " + branch + ".";
  } else {
    encText = String(status || "-");
  }
  return encText;
}

/* ========= SR Detail ========= */
function addSRDetailRow(){
  var tpl = $("#tplSRDetail");
  if(!tpl){ console.warn("tplSRDetail missing"); return document.createElement("div"); }
  var row = tpl.content.firstElementChild.cloneNode(true);
  var btn = $(".btnRemoveSR", row);
  if(btn) btn.addEventListener("click", function(){ row.remove(); saveState(); });
  var container = $("#srDetailContainer");
  if(container) container.appendChild(row);
  return row;
}

/* ========= Search Amount calc ========= */
function recalcSearchAmount(){
  var y1 = parseInt($("#srYearFrom") ? $("#srYearFrom").value : "0", 10);
  var y2 = parseInt($("#srYearTo") ? $("#srYearTo").value : "0", 10);
  var amtEl = $("#searchAmount");
  if(y1>0 && y2>=y1){
    var amt = (y2 - y1 + 1) * 50;
    if(amtEl) amtEl.value = String(amt);
  } else { if(amtEl) amtEl.value = ""; }
  saveState();
}

/* ========= Property formatter used by Preview + Excel ========= */
/* (unchanged - uses asHtml flag) */
function formatPropertyDetails(props, asHtml){
  props = props || [];
  asHtml = !!asHtml;
  var out = "";
  var finalTotal=0, finalCommand=0, finalUncommand=0;
  var finalKhala=0, finalRasta=0, finalKuan=0, finalNahari=0, finalNali1=0, finalNali2=0, finalBarani1=0;
  var shareFlag=false;
  var multiProps = props.length > 1;

  props.forEach(function(ps){
    var chak = ps.fixed?.chak || "-";
    var samvat = ps.fixed?.samvat || "-";
    var acnoRaw = ps.fixed?.acno || "";
    var acnos = acnoRaw.split ? acnoRaw.split(",").map(function(a){ return a.trim(); }).filter(Boolean) : [];
    var acStr = acnos.length === 0 ? "A/c No. -" : (acnos.length === 1 ? "A/c No. " + acnos[0] : "A/c No. " + acnos.slice(0,-1).join(", ") + " & " + acnos.slice(-1));

    var headerLine = "Chak " + chak + " Jamabandi Samvat " + samvat + " " + acStr;
    out += asHtml ? ("<div><strong>" + headerLine + "</strong></div>") : (headerLine + "\n");

    // group S.No / Murba / K.Nos
    var rows = ps.rows || [];
    var groups = {};
    var order = [];
    var lastS = null, lastM = null;
    rows.forEach(function(r){
      var rawS = r.sno && r.sno !== "-" ? r.sno : null;
      var rawM = r.murba && r.murba !== "-" ? r.murba : null;
      var sno = rawS !== null ? rawS : lastS;
      var murba = rawM !== null ? rawM : lastM;
      if(rawS !== null) lastS = rawS;
      if(rawM !== null) lastM = rawM;
      if(!r.kno) return;
      var key = (sno || "") + "|" + (murba || "");
      if(!groups[key]){ groups[key] = { sno: sno, murba: murba, knos: [] }; order.push(key); }
      var areaPart = r.area ? ("/" + r.area) : "";
      groups[key].knos.push(r.kno + areaPart);
    });
    order.forEach(function(k){
      var g = groups[k];
      if(!g || !g.knos.length) return;
      var murbaTxt = g.murba ? " (" + g.murba + ")" : "";
      var snoTxt = g.sno ? ("S.No: " + g.sno + murbaTxt) : ("S.No:" + murbaTxt);
      var line = snoTxt + " K.No " + g.knos.join(", ") + " Hect.";
      out += asHtml ? ("<div>" + line + "</div>") : (line + "\n");
    });

    // totals for this property
    var tH = num(ps.totals?.total), tC = num(ps.totals?.command), tU = num(ps.totals?.uncommand);
    var tKhala = num(ps.totals?.khala), tRasta = num(ps.totals?.rasta), tKuan = num(ps.totals?.kuan);
    var tNahari = num(ps.totals?.nahari), tNali1 = num(ps.totals?.nali1), tNali2 = num(ps.totals?.nali2), tBarani1 = num(ps.totals?.barani1);

    var tParts = [];
    if(tH) tParts.push("Total Land = " + tH.toFixed(4) + " Hect.");
    if(tC) tParts.push("Command = " + tC.toFixed(4) + " Hect.");
    if(tU) tParts.push("Uncommand = " + tU.toFixed(4) + " Hect.");
    if(tKhala) tParts.push("G.M. Khala = " + tKhala.toFixed(4) + " Hect.");
    if(tRasta) tParts.push("G.M. Rasta = " + tRasta.toFixed(4) + " Hect.");
    if(tKuan) tParts.push("G.M. Kuan = " + tKuan.toFixed(4) + " Hect.");
    if(tNahari) tParts.push("Nahari = " + tNahari.toFixed(4) + " Hect.");
    if(tNali1) tParts.push("Nali 1st = " + tNali1.toFixed(4) + " Hect.");
    if(tNali2) tParts.push("Nali 2nd = " + tNali2.toFixed(4) + " Hect.");
    if(tBarani1) tParts.push("Barani 1st = " + tBarani1.toFixed(4) + " Hect.");
    if(tParts.length){
      var tLine = tParts.join(", ");
      out += asHtml ? ("<div>" + tLine + "</div>") : (tLine + "\n");
    }

    // share of borrower if present
    var shareStr = (ps.borrowerShare || "").trim();
    var effective = tH;
    if(shareStr){
      var f = parseShareFactor(shareStr);
      if(f>0 && tH){
        shareFlag = true;
        effective = tH * f;
        var sC = tC * f, sU = tU * f;
        var sKhala = tKhala * f, sRasta = tRasta * f, sKuan = tKuan * f;
        var sNahari = tNahari * f, sNali1 = tNali1 * f, sNali2 = tNali2 * f, sBarani1 = tBarani1 * f;
        var parts = ["Share of Borrower = " + shareStr + " = " + effective.toFixed(4) + " Hect."];
        if(sC) parts.push("Command = " + sC.toFixed(4) + " Hect.");
        if(sU) parts.push("Uncommand = " + sU.toFixed(4) + " Hect.");
        if(sKhala) parts.push("G.M. Khala = " + sKhala.toFixed(4) + " Hect.");
        if(sRasta) parts.push("G.M. Rasta = " + sRasta.toFixed(4) + " Hect.");
        if(sKuan) parts.push("G.M. Kuan = " + sKuan.toFixed(4) + " Hect.");
        if(sNahari) parts.push("Nahari = " + sNahari.toFixed(4) + " Hect.");
        if(sNali1) parts.push("Nali 1st = " + sNali1.toFixed(4) + " Hect.");
        if(sNali2) parts.push("Nali 2nd = " + sNali2.toFixed(4) + " Hect.");
        if(sBarani1) parts.push("Barani 1st = " + sBarani1.toFixed(4) + " Hect.");
        out += asHtml ? ("<div>" + parts.join(", ") + "</div>") : (parts.join(", ") + "\n");

        finalTotal += effective; finalCommand += sC; finalUncommand += sU;
        finalKhala += sKhala; finalRasta += sRasta; finalKuan += sKuan;
        finalNahari += sNahari; finalNali1 += sNali1; finalNali2 += sNali2; finalBarani1 += sBarani1;
      }
    } else {
      finalTotal += tH; finalCommand += tC; finalUncommand += tU;
      finalKhala += tKhala; finalRasta += tRasta; finalKuan += tKuan;
      finalNahari += tNahari; finalNali1 += tNali1; finalNali2 += tNali2; finalBarani1 += tBarani1;
    }

    // single-property -> show acres now
    if(!multiProps && effective > 0){
      var a = hectaresToAcresText(effective);
      var aLine = "Area in Acres = " + a.text;
      out += asHtml ? ("<div>" + aLine + "</div>") : (aLine + "\n");
    }
  }); // end props loop

  // multi-property -> grand total only
  if(multiProps){
    if(shareFlag){
      var sLine = "Grand Total Share of Borrower = " + finalTotal.toFixed(4) + " Hect.";
      if(finalCommand) sLine += ", Command = " + finalCommand.toFixed(4) + " Hect.";
      if(finalUncommand) sLine += ", Uncommand = " + finalUncommand.toFixed(4) + " Hect.";
      if(finalKhala) sLine += ", G.M. Khala = " + finalKhala.toFixed(4) + " Hect.";
      if(finalRasta) sLine += ", G.M. Rasta = " + finalRasta.toFixed(4) + " Hect.";
      if(finalKuan) sLine += ", G.M. Kuan = " + finalKuan.toFixed(4) + " Hect.";
      if(finalNahari) sLine += ", Nahari = " + finalNahari.toFixed(4) + " Hect.";
      if(finalNali1) sLine += ", Nali 1st = " + finalNali1.toFixed(4) + " Hect.";
      if(finalNali2) sLine += ", Nali 2nd = " + finalNali2.toFixed(4) + " Hect.";
      if(finalBarani1) sLine += ", Barani 1st = " + finalBarani1.toFixed(4) + " Hect.";
      out += asHtml ? ("<div><strong>" + sLine + "</strong></div>") : (sLine + "\n");
      var a = hectaresToAcresText(finalTotal);
      var aLine = "Area in Acres = " + a.text;
      out += asHtml ? ("<div><strong>" + aLine + "</strong></div>") : (aLine + "\n");
    } else {
      var gLine = "Grand Total Land of Borrower = " + finalTotal.toFixed(4) + " Hect.";
      if(finalCommand) gLine += ", Command = " + finalCommand.toFixed(4) + " Hect.";
      if(finalUncommand) gLine += ", Uncommand = " + finalUncommand.toFixed(4) + " Hect.";
      if(finalKhala) gLine += ", G.M. Khala = " + finalKhala.toFixed(4) + " Hect.";
      if(finalRasta) gLine += ", G.M. Rasta = " + finalRasta.toFixed(4) + " Hect.";
      if(finalKuan) gLine += ", G.M. Kuan = " + finalKuan.toFixed(4) + " Hect.";
      if(finalNahari) gLine += ", Nahari = " + finalNahari.toFixed(4) + " Hect.";
      if(finalNali1) gLine += ", Nali 1st = " + finalNali1.toFixed(4) + " Hect.";
      if(finalNali2) gLine += ", Nali 2nd = " + finalNali2.toFixed(4) + " Hect.";
      if(finalBarani1) gLine += ", Barani 1st = " + finalBarani1.toFixed(4) + " Hect.";
      out += asHtml ? ("<div><strong>" + gLine + "</strong></div>") : (gLine + "\n");
      var a = hectaresToAcresText(finalTotal);
      var aLine = "Area in Acres = " + a.text;
      out += asHtml ? ("<div><strong>" + aLine + "</strong></div>") : (aLine + "\n");
    }
  }

  return out;
}

/* ========= Preview ========= */
function renderPreview(){
  var st = collectState();
  var html = "";

  // ===== Step 1: बैंक एवं उधारकर्ता विवरण =====
  html += `<h3>बैंक एवं उधारकर्ता विवरण</h3>`;
  html += `<div><strong>Sr. No:</strong> ${st.srNo || "-"} | <strong>Date:</strong> ${st.bankDate || "-"}</div>`;
  html += `<div><strong>Bank Name:</strong> ${st.bankName || "-"} 
| <strong>Branch (Hindi):</strong> <span class="hindi-text">${st.branchHindi || "-"}</span> 
| <strong>Branch (English):</strong> ${st.branchEnglish || "-"}</div>`;

  html += `<div><strong>Borrower (Hindi):</strong> <span class="hindi-text">${st.nameHindi || "-"}</span> 
| <strong>Borrower (English):</strong> ${st.nameEnglish || "-"}</div>`;

  if(st.propertySets && st.propertySets.length){
    html += "<hr><h3>Property Details</h3>";
    html += formatPropertyDetails(st.propertySets, true);
  }
  html += "<hr><h3>Boundaries</h3>";
  html += "<div>North: " + (st.boundaries?.north || "-") + "</div>";
  html += "<div>South: " + (st.boundaries?.south || "-") + "</div>";
  html += "<div>East: " + (st.boundaries?.east || "-") + "</div>";
  html += "<div>West: " + (st.boundaries?.west || "-") + "</div>";

  // ===== Encumbrance (use common helper so Excel + Preview match) =====
  html += "<hr><h3>Encumbrance</h3>";
  html += "<div>" + buildEncumbranceText(st) + "</div>";

  if(st.srDetails && st.srDetails.length){
    html += "<hr><h3>SR Details</h3>";
    st.srDetails.forEach(function(d,i){ html += "<div>" + (i+1) + ". Office: " + (d.office||"-") + ", Receipt: " + (d.receipt||"-") + ", GRN: " + (d.grn||"-") + "</div>"; });
  }

  html += "<hr><h3>Search Details</h3>";
  html += "<div>Year From: " + (st.searchDetail?.yearFrom || "-") + " | Year To: " + (st.searchDetail?.yearTo || "-") + "</div>";
  html += "<div>Amount: " + (st.searchDetail?.amount || "-") + " | Gift/Other: " + (st.searchDetail?.giftOther || "-") + " | DLC Rate: " + (st.searchDetail?.dlcRate || "-") + "</div>";

  html += "<hr><h3>Documents</h3>";
  html += "<div>Jamabandi: " + (st.documents?.jamabandi || "-") + ", Girdawari: " + (st.documents?.girdawari || "-") + ", Map: " + (st.documents?.map || "-") + ", Samvat: " + (st.documents?.samvat || "-") + ", Other: " + (st.documents?.other || "-") + "</div>";

  // ===== Advocate Details: use saved state (st.advocates) so preview matches saved data =====
  if(Array.isArray(st.advocates) && st.advocates.length){
    html += "<hr><h3>Advocate Details</h3>";
    st.advocates.forEach(function(a,i){
      var name = a.name || "-";
      var mobile = a.mobile || "-";
      var chamber = a.chamber || "-";
      var other = a.other || "";
      html += "<div>" + (i+1) + ". " + name + " | Mobile: " + mobile + " | Chamber: " + chamber + (other ? " | Other: " + other : "") + "</div>";
    });
  }

  var box = $("#previewBox");
  if(box) box.innerHTML = html; else console.warn("#previewBox not found");
}

/* ========= Excel builder ========= */
function _buildExcelRowFromState(st, idx){
  st = st || collectState();
  var propertyDetails = formatPropertyDetails(st.propertySets || [], false).trim();

  var boundaries = "North - " + (st.boundaries?.north || "-") + 
                   ", South - " + (st.boundaries?.south || "-") + 
                   ", East - " + (st.boundaries?.east || "-") + 
                   ", West - " + (st.boundaries?.west || "-");

  // ---- Location & Primary Chak ----
  var locParts = [];
  var chakSet = new Set();
  if(st.propertySets && st.propertySets.length){
    var groups = {};
    st.propertySets.forEach(function(ps){
      var chak = ps.fixed?.chak || "-";
      chakSet.add(chak);
      var tehsil = ps.fixed?.tehsil || "-";
      var distt = ps.fixed?.distt || "-";
      var key = tehsil + "|" + distt;
      if(!groups[key]) groups[key] = { tehsil: tehsil, distt: distt, chaks: new Set() };
      groups[key].chaks.add(chak);
    });
    Object.values(groups).forEach(function(g){
      var chaks = Array.from(g.chaks);
      var chakStr = chaks.length === 1 ? "Chak " + chaks[0] : "Chak " + chaks.join(" & Chak ");
      locParts.push(chakStr + " Tehsil " + g.tehsil + ", Distt. " + g.distt);
    });
  }
  var location = locParts.join(" & ");
  var primaryChak = "-";
  if(chakSet.size === 1){ primaryChak = Array.from(chakSet)[0]; }
  else if(chakSet.size > 1){ primaryChak = Array.from(chakSet).join(" & Chak "); }

  // ---- SR / Inspection ----
  var srOffice = (st.srDetails || []).map(d=>d.office||"-").filter(Boolean).join(", ");
  var inspectionYears = (st.searchDetail?.yearFrom && st.searchDetail?.yearTo) 
      ? (st.searchDetail.yearFrom + "-" + st.searchDetail.yearTo) : "-";
  var srReceipts = (st.srDetails || []).map(function(d){ 
    return "Search Inspection Fee,Receipt " + (d.office||"-") + " S.R. No. " + (d.receipt||"-"); 
  }).join("\n");

  // ---- Jamabandi, Girdawari, Map ----
  var jamabandiDetail = "";
  if(st.propertySets && st.propertySets.length){
    jamabandiDetail = st.propertySets.map(function(ps){
      var samvat = ps.fixed?.samvat || "-";
      var chak = ps.fixed?.chak || "-";
      var acnos = (ps.fixed?.acno || "").split(",").map(a=>a.trim()).filter(Boolean);
      var acStr = acnos.length === 0 ? "A/c No. -" : (acnos.length === 1 ? "A/c No. "+acnos[0] : "A/c No. "+acnos.slice(0,-1).join(", ")+" & "+acnos.slice(-1));
      return "Jamabandi Samvat " + samvat + " of Chak " + chak + " " + acStr;
    }).join(" & ") + " Dt. " + (st.documents?.jamabandi || "-");
  }

  var girdawariDetail = "";
  if(st.propertySets && st.propertySets.length){
    girdawariDetail = st.propertySets.map(function(ps){
      var samvat = st.documents?.samvat || "-";
      var chak = ps.fixed?.chak || "-";
      var acnos = (ps.fixed?.acno || "").split(",").map(a=>a.trim()).filter(Boolean);
      var acStr = acnos.length === 0 ? "A/c No. -" : (acnos.length === 1 ? "A/c No. "+acnos[0] : "A/c No. "+acnos.slice(0,-1).join(", ")+" & "+acnos.slice(-1));
      return "Girdawari Samvat " + samvat + " of Chak " + chak + " " + acStr;
    }).join(" & ") + " Dt. " + (st.documents?.girdawari || "-");
  }

  var mapOfLand = "";
  if(st.propertySets && st.propertySets.length){
    mapOfLand = st.propertySets.map(function(ps){
      var chak = ps.fixed?.chak || "-";
      var acnos = (ps.fixed?.acno || "").split(",").map(a=>a.trim()).filter(Boolean);
      var acStr = acnos.length === 0 ? "A/c No. -" : (acnos.length === 1 ? "A/c No. "+acnos[0] : "A/c No. "+acnos.slice(0,-1).join(", ")+" & "+acnos.slice(-1));
      return "Map of Chak " + chak + " " + acStr;
    }).join(" & ") + " Dt. " + (st.documents?.map || "-");
  }

  // ---- Encumbrance full sentence ----
  var encText = "-";
  if(st.encumbrance && st.encumbrance.status){
    var bank = st.bankName || "_____";
    var branch = st.branchEnglish || st.branchHindi || "_____";
    var status = st.encumbrance.status;
    if(status === "mortgagedFull" || status === "mortgaged" || status === "1"){
      encText = "The Land/Share of Borrower (s) is already mortgaged in favour of " + bank + " Branch " + branch + ".";
    } else if(status === "free" || status === "2"){
      encText = "The Land/Share of Borrower (s) is free from all encumbrances.";
    } else if(status === "mortgagedPartial" || status === "3"){
      var manual = st.encumbrance.manual || "{without Land of Chak/Stone/Killa../Area}";
      encText = "The Land / Share of Borrower " + manual + " is already mortgaged in favour of " + bank + " Branch " + branch + ".";
    }
  }

  // ---- Total Land (Hect & Acres) ----
  var totalLandShort = "-";
  if(st.propertySets && st.propertySets.length){
    if(st.propertySets.length === 1){ 
      // Single Property
      var ps = st.propertySets[0];
      var tH = parseFloat(ps.totals?.total) || 0;
      var shareStr = (ps.borrowerShare || "").trim();
      if(shareStr){
        var f = parseShareFactor(shareStr);
        if(f>0 && tH>0){
          var eff = tH * f;
          totalLandShort = eff.toFixed(4) + " Hect., " + hectaresToAcresText(eff).acres.toFixed(3) + " Acres";
        }
      } else if(tH>0){
        totalLandShort = tH.toFixed(4) + " Hect., " + hectaresToAcresText(tH).acres.toFixed(3) + " Acres";
      }
    } else {
      // Multi Property
      var grandTotal=0, grandShare=0, shareFlag=false;
      st.propertySets.forEach(function(ps){
        var tH = parseFloat(ps.totals?.total) || 0;
        var shareStr = (ps.borrowerShare || "").trim();
        if(shareStr){
          var f = parseShareFactor(shareStr);
          if(f>0 && tH>0){ grandShare += tH * f; shareFlag=true; }
        }
        grandTotal += tH;
      });
      if(shareFlag && grandShare>0){
        totalLandShort = grandShare.toFixed(4) + " Hect., " + hectaresToAcresText(grandShare).acres.toFixed(3) + " Acres";
      } else if(grandTotal>0){
        totalLandShort = grandTotal.toFixed(4) + " Hect., " + hectaresToAcresText(grandTotal).acres.toFixed(3) + " Acres";
      }
    }
  }

  // ---- Advocate ----
  var advName="-", advMobile="-", advChamber="-";
  if(Array.isArray(st.advocates) && st.advocates.length){
    advName = st.advocates[0].name || "-";
    advMobile = st.advocates[0].mobile || "-";
    advChamber = st.advocates[0].chamber || "-";
  }

  // ---- Headers + Row ----
  var headers = [
    "Sr. No","Date","Bank Name","Branch (Hindi)","Branch (English)","Borrower (Hindi)","Borrower (English)",
    "Property Detail","Total Land (Hect & Acres)","Encumbrance Status","Boundaries","Location","S.R. Offices",
    "Inspection Years","SR Receipts","Jamabandi Detail","Girdawari Detail","Map of Land",
    "Gift/Other","DLC Rate","Primary Chak","Advocate Name","Advocate Mobile","Advocate Chamber"
  ];

  var row = [
    st.srNo || String(idx? idx+1 : "1"),
    st.bankDate || "",
    st.bankName || "",
    st.branchHindi || "",
    st.branchEnglish || "",
    st.nameHindi || "",
    st.nameEnglish || "",
    propertyDetails,
    totalLandShort,
    encText,
    boundaries,
    location,
    srOffice,
    inspectionYears,
    srReceipts,
    jamabandiDetail,
    girdawariDetail,
    mapOfLand,
    st.searchDetail?.giftOther || "-",
    st.searchDetail?.dlcRate ? ("DLC Rate = " + st.searchDetail.dlcRate + "/- per Hect.") : "-",
    primaryChak,
    advName,
    advMobile,
    advChamber
  ];

  return { headers: headers, row: row };

// 🔥 upfront declaration for Excel export column
  var totalLandShort = "-";

  // ===== FIX: Total Land (Hect & Acres) column logic (Excel only) =====
  try {
    var props = Array.isArray(st.propertySets) ? st.propertySets : [];
    var multiProps = props.length > 1;
    var shareFlag = props.some(function(ps){ return String(ps.borrowerShare||'').trim() !== ''; });

    var sumH = 0;
    props.forEach(function(ps){
      var tH = parseFloat(ps && ps.totals && ps.totals.total) || 0;
      var shareStr = String(ps && ps.borrowerShare || '').trim();
      if(shareStr){
        var f = parseShareFactor(shareStr);
        sumH += tH * f;
      } else {
        sumH += tH;
      }
    });

    if(!multiProps && props.length === 1){
      var tH1 = parseFloat(props[0].totals && props[0].totals.total) || 0;
      var shareStr1 = String(props[0].borrowerShare || '').trim();
      if(tH1 > 0){
        if(shareStr1){
          var f1 = parseShareFactor(shareStr1);
          var eff1 = tH1 * f1;
          totalLandShort = eff1.toFixed(4) + " Hect., " +
                           hectaresToAcresText(eff1).acres.toFixed(3) + " Acres";
        } else {
          totalLandShort = tH1.toFixed(4) + " Hect., " +
                           hectaresToAcresText(tH1).acres.toFixed(3) + " Acres";
        }
      }
    } else if(multiProps){
      if(shareFlag){
        totalLandShort = sumH.toFixed(4) + " Hect., " +
                         hectaresToAcresText(sumH).acres.toFixed(3) + " Acres";
      } else {
        totalLandShort = sumH.toFixed(4) + " Hect., " +
                         hectaresToAcresText(sumH).acres.toFixed(3) + " Acres";
      }
    }
  } catch(e){ totalLandShort = "-"; }

  // ===== Post-processing columns for Excel Export =====
  if(typeof idx !== 'undefined'){ srNo = idx+1; } else { srNo = ''; }

  // Encumbrance fixed text
  var encumbranceText = "The Land/Share of Borrower (s) is already mortgaged in favour of State Bank of India Branch Silwala Khurd.";

  // Jamabandi, Girdawari, Map - join with & and put Date only at the end
  function joinWithDate(entries, date){
    if(!entries || !entries.length) return '';
    var core = entries.join(' & ');
    return core + (date ? ' Dt. ' + date : '');
  }

  var jamabandiEntries = (st.jamabandiDetails||[]).map(x=>x.text||x);
  var girdawariEntries = (st.girdawariDetails||[]).map(x=>x.text||x);
  var mapEntries = (st.mapDetails||[]).map(x=>x.text||x);
  var commonDate = st.date || '';

  var jamabandiText = joinWithDate(jamabandiEntries, commonDate);
  var girdawariText = joinWithDate(girdawariEntries, commonDate);
  var mapText = joinWithDate(mapEntries, commonDate);

  // Primary Chak join logic
  var chakList = (st.primaryChaks||[]).map(x=>x.name||x);
  var primaryChak = '';
  if(chakList.length===1) primaryChak = chakList[0];
  else if(chakList.length>1) primaryChak = chakList.join(' & ');

}

/* ========= Export (uses XLSX) ========= */
function exportExcel(){
  if(typeof XLSX === "undefined"){ alert("XLSX library not loaded. Please include xlsx.full.min.js"); return; }
  var st = collectState();
  var built = _buildExcelRowFromState(st);
  var wb = XLSX.utils.book_new();
  var ws = XLSX.utils.aoa_to_sheet([built.headers, built.row]);
  XLSX.utils.book_append_sheet(wb, ws, "TIR Report");
  XLSX.writeFile(wb, "SBI_TIR_Database.xlsx");
}

function exportAllEntries(){
  if(typeof XLSX === "undefined"){ alert("XLSX library not loaded. Please include xlsx.full.min.js"); return; }
  var list = [];
  try{ list = JSON.parse(localStorage.getItem(LS_ENTRIES) || "[]"); }catch(e){ list = []; }
  if(!Array.isArray(list) || list.length === 0){ alert("कोई saved entry नहीं मिली. 'Final Submit' से पहले entries सेव करें."); return; }
  var wb = XLSX.utils.book_new();
  var headers = null;
  var rows = [];
  list.forEach(function(st, idx){
    var built = _buildExcelRowFromState(st);
    if(!headers) headers = built.headers;
    if(!st.srNo) built.row[0] = String(idx+1);
    rows.push(built.row);
  });
  var ws = XLSX.utils.aoa_to_sheet([headers].concat(rows));
  XLSX.utils.book_append_sheet(wb, ws, "All Entries");
  XLSX.writeFile(wb, "SBI_TIR_All_Entries.xlsx");
}

/* ========= Events bind ========= */
function bindEvents(){
  $$(".stepper .step-node").forEach(function(btn){ btn.addEventListener("click", function(){ goTo(btn.dataset.goto); }); });
  $$(".goBack").forEach(function(btn){ btn.addEventListener("click", function(){ goTo(btn.dataset.goto); }); });

  var el = $("#s1SaveNext"); if(el) el.addEventListener("click", function(){ saveState(); goTo(2); });
  el = $("#s2SaveNext"); if(el) el.addEventListener("click", function(){ saveState(); renderPreview(); goTo(3); });

  el = $("#btnAddPropertySet"); if(el) el.addEventListener("click", function(){ var node = addPropertySet(); saveState(); if(node) node.scrollIntoView({behavior:"smooth", block:"start"}); });

  document.addEventListener("change", function(e){ if(e.target && e.target.matches && e.target.matches("input, select")) saveState(); });

  el = $("#exportBtn"); if(el) el.addEventListener("click", exportExcel);
  el = $("#exportAllBtn"); if(el) el.addEventListener("click", exportAllEntries);
  el = $("#clearAllBtn"); if(el) el.addEventListener("click", function(){ if(confirm("सारा लोकल डेटा साफ़ कर दिया जाएगा.")){ localStorage.removeItem(LS_FORM); location.reload(); } });
  el = $("#resetSrBtn"); if(el) el.addEventListener("click", function(){ if(confirm("सारा लोकल डेटा साफ़ कर दिया जाएगा.")){ localStorage.removeItem(LS_FORM); location.reload(); } });

  el = $("#encumbranceStatus"); if(el) el.addEventListener("change", function(){ toggleEncumbranceManual(); saveState(); });
  el = $("#encumbranceManual"); if(el) el.addEventListener("input", saveState);

  el = $("#btnAddSRDetail"); if(el) el.addEventListener("click", function(){ addSRDetailRow(); saveState(); });

  el = $("#srYearFrom"); if(el) el.addEventListener("input", recalcSearchAmount);
  el = $("#srYearTo"); if(el) el.addEventListener("input", recalcSearchAmount);

  var finalBtn = document.getElementById("finalSubmitBtn");
  if(finalBtn){
    finalBtn.addEventListener("click", function(){
      var entry = collectState();
      if(!entry){ alert("कृपया आवश्यक फ़ील्ड भरें"); return; }
      var allData = JSON.parse(localStorage.getItem(LS_ENTRIES) || "[]");
      allData.push(entry);
      localStorage.setItem(LS_ENTRIES, JSON.stringify(allData));
      alert("✅ Entry Final Submit हो गई");
      document.querySelector("form") && document.querySelector("form").reset();
      localStorage.removeItem(LS_FORM);
      location.reload();
    });
  } else { console.warn("#finalSubmitBtn not found"); }

  // advocate add (unchanged)
  var advocateContainer = $("#advocateContainer");
  var btnAddAdvocate = $("#btnAddAdvocate");
  var tplAdvRow = $("#tplAdvocateRow");
  if(advocateContainer && btnAddAdvocate && tplAdvRow){
    btnAddAdvocate.addEventListener("click", function(){
      var clone = tplAdvRow.content.cloneNode(true);
      var row = clone.querySelector(".advocate-row");
      advocateContainer.appendChild(clone);
      var sel = row.querySelector("#advocateName");
      var mobile = row.querySelector("#advocateMobile");
      var chamber = row.querySelector("#advocateChamber");
      var otherWrap = row.querySelector("#advocateOtherWrap");
      var otherInput = row.querySelector("#advocateOtherInput");
      var rem = row.querySelector(".btnRemoveAdvocate");
      var advocates = {"Mahesh Kumar Gaur": {mobile:"9462949117", chamber:"Chamber No. 2"}, "Bhanwar Lal Gaur": {mobile:"9414434217", chamber:"Chamber No. 2"}};
      if(sel) sel.addEventListener("change", function(){
        var v = sel.value;
        if(advocates[v]){ if(mobile) mobile.value = advocates[v].mobile; if(chamber) chamber.value = advocates[v].chamber; if(otherWrap) otherWrap.style.display="none"; if(otherInput) otherInput.value=""; }
        else if(v === "other"){ if(otherWrap) otherWrap.style.display="block"; if(mobile) mobile.value = ""; if(chamber) chamber.value = ""; }
        else { if(otherWrap) otherWrap.style.display="none"; if(mobile) mobile.value = ""; if(chamber) chamber.value = ""; if(otherInput) otherInput.value=""; }
      });
      if(rem) rem.addEventListener("click", function(){ row.remove(); });
    });
    // add one default row
    btnAddAdvocate.click();
  }
}

/* ========= Init ========= */
document.addEventListener("DOMContentLoaded", function(){
  try{
    bindEvents();
    loadState();
    if($$("#propertySetContainer .property-set").length === 0) addPropertySet();
    if($("#srDetailContainer") && $("#srDetailContainer").children.length === 0) addSRDetailRow();
    recalcSearchAmount();
    recalcGrandTotals();
  }catch(e){
    console.error("Initialization error:", e);
  }
});
