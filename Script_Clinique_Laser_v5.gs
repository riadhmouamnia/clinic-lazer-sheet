/** Gestion clinique laser — v5
 * Le classeur ne contient plus que deux onglets :
 *   « Registre patients »    : la base de données (une ligne = un rendez-vous)
 *   « Liste des lazéristes » : les noms proposés dans l'application (à partir de A5)
 * Tout le reste se fait dans l'application (menu Clinique laser → Ouvrir l'application).
 *
 * Installation : importez Clinique_Laser_v5.xlsx dans Google Drive (ouvrir avec Google Sheets),
 * collez ce code dans Extensions > Apps Script (Code.gs), ajoutez un fichier HTML nommé « App »
 * avec le contenu de App.html, enregistrez, rechargez le classeur, puis Clinique laser → Configurer le registre.
 */
const REGISTRE='Registre patients', LAZERISTES='Liste des lazéristes';
const ENTETES=['Terminé ?','Date d\'enregistrement','Nom','Prénom','Région / Wilaya','Téléphone','Heure d\'arrivée','Date du rendez-vous','Heure du rendez-vous','Lazériste','Paramètres / Type de laser','Séances','Prix (DA)','Remarques'];
// Formats des colonnes B..N. Le texte « @ » garde le 0 du téléphone et n'est jamais interprété comme date ou formule.
const FORMATS=['dd/MM/yyyy','@','@','@','@','HH:mm','dd/MM/yyyy','HH:mm','@','@','0','#,##0.00 "DA"','@'];

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Clinique laser')
    .addItem('Ouvrir l’application','ouvrirApplication')
    .addSeparator()
    .addItem('Configurer le registre','configurerRegistre')
    .addToUi();
}

/* ───────────── Mise en place du classeur ───────────── */
// Applique le thème du registre ; peut être relancé sans risque, les données ne sont pas modifiées.
function configurerRegistre(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(REGISTRE);
  if(!sh||!ss.getSheetByName(LAZERISTES)) throw new Error('Les onglets « '+REGISTRE+' » et « '+LAZERISTES+' » sont nécessaires.');
  if(sh.getMaxRows()<1001) sh.insertRowsAfter(sh.getMaxRows(),1001-sh.getMaxRows());
  if(sh.getMaxColumns()>14) sh.deleteColumns(15,sh.getMaxColumns()-14);
  const n=sh.getMaxRows()-1, corps=sh.getRange(2,1,n,14);
  sh.getRange('A1:N1').setValues([ENTETES]).setBackground('#203864').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1,44);
  [80,110,140,140,140,125,90,110,90,140,200,75,115,320].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  // Couleurs portées uniquement par la mise en forme conditionnelle : aucune couleur figée dans les cellules.
  corps.setBackground(null).setFontColor('#1B2433').setFontFamily('Arial').setFontSize(10).setFontWeight('normal')
    .setVerticalAlignment('middle').setWrap(false).clearDataValidations()
    .setBorder(false,false,true,false,false,true,'#E3E8F0',SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(2,3,n,1).setFontWeight('bold');
  sh.getRange(2,2,n,13).setNumberFormats(Array(n).fill(FORMATS));
  [['A','center'],['B','center'],['G','center'],['H','center'],['I','center'],['L','center'],['M','right']]
    .forEach(([c,a])=>sh.getRange(c+'2:'+c+(n+1)).setHorizontalAlignment(a));
  sh.setRowHeights(2,n,28);
  sh.getRange(2,1,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  reglesRegistre_(sh);
  sh.setFrozenRows(1); sh.setFrozenColumns(4); sh.setHiddenGridlines(true); sh.setTabColor('#203864');
  if(!sh.getFilter()) sh.getRange(1,1,sh.getMaxRows(),14).createFilter();
  SpreadsheetApp.getUi().alert('Configuration terminée','Le registre est mis en forme. Utilisez Clinique laser → Ouvrir l’application pour enregistrer et gérer les patients.',SpreadsheetApp.getUi().ButtonSet.OK);
}
// Première règle qui correspond = seule appliquée : ligne verte si terminé, case rouge si en attente, sinon une ligne sur deux teintée.
function reglesRegistre_(sh){
  const n=sh.getMaxRows(), tout=sh.getRange('A2:N'+n), regle=SpreadsheetApp.newConditionalFormatRule;
  sh.setConditionalFormatRules([
    regle().whenFormulaSatisfied('=$A2').setBackground('#E3F3E9').setFontColor('#1E5631').setRanges([tout]).build(),
    regle().whenFormulaSatisfied('=AND(NOT($A2),$C2<>"")').setBackground('#FBE4E2').setRanges([sh.getRange('A2:A'+n)]).build(),
    regle().whenFormulaSatisfied('=AND(ISEVEN(ROW()),$C2<>"")').setBackground('#F7F9FC').setRanges([tout]).build()
  ]);
}

/* ───────────── Lecture du registre ───────────── */
// Toutes les lignes avec un patient. y/m/d (rendez-vous) et ey/em/ed (enregistrement) valent 0 si la date est vide ou illisible.
function lirePatients_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(REGISTRE), tz=ss.getSpreadsheetTimeZone();
  const n=sh.getLastRow()-1; if(n<1) return [];
  const rng=sh.getRange(2,1,n,14), vals=rng.getValues(), txt=rng.getDisplayValues(), out=[], p2=x=>('0'+x).slice(-2);
  vals.forEach((v,i)=>{
    const t=txt[i], nom=String(v[2]).trim(), prenom=String(v[3]).trim();
    if(!nom&&!prenom) return;
    const j=jourDe_(v[7],t[7],tz)||[0,0,0], e=jourDe_(v[1],t[1],tz)||[0,0,0];
    out.push({ligne:i+2,emp:empreinte_(t),enreg:t[1],ey:e[0],em:e[1],ed:e[2],nom:nom,prenom:prenom,region:String(v[4]).trim(),tel:t[5],arrivee:heureDe_(v[6],t[6],tz),
      y:j[0],m:j[1],d:j[2],date:j[0]?j[0]+'-'+p2(j[1])+'-'+p2(j[2]):'',dateTxt:t[7].trim(),heure:heureDe_(v[8],t[8],tz),
      laz:String(v[9]).trim(),params:String(v[10]).trim(),seances:t[11],seancesNum:typeof v[11]==='number'?v[11]:'',
      prix:t[12],prixNum:typeof v[12]==='number'?v[12]:'',remarques:String(v[13]).trim(),termine:v[0]===true});
  });
  return out;
}
// Le texte affiché (dd/MM/yyyy, HH:mm) est prioritaire : les heures stockées en 1899 peuvent être décalées par le fuseau.
function jourDe_(val,txt,tz){
  const m=String(txt).trim().match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if(m) return [+m[3],+m[2],+m[1]];
  return val instanceof Date&&!isNaN(val)?Utilities.formatDate(val,tz,'yyyy-M-d').split('-').map(Number):null;
}
function heureDe_(val,txt,tz){
  const m=String(txt).trim().match(/^(\d{1,2})[:hH.](\d{2})/);
  if(m) return ('0'+m[1]).slice(-2)+':'+m[2];
  return val instanceof Date&&!isNaN(val)?Utilities.formatDate(val,tz,'HH:mm'):'';
}
function aujourdhui_(){ return Utilities.formatDate(new Date(),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),'yyyy-M-d').split('-').map(Number); }

/* ───────────── Application (fenêtre ou page web) ───────────── */
function ouvrirApplication(){
  SpreadsheetApp.getUi().showModelessDialog(page_('fenetre').setWidth(1240).setHeight(800),'Clinique laser');
}
// Page web : Déployer > Nouveau déploiement > Application Web.
function doGet(){
  return page_('web').setTitle('Clinique laser').addMetaTag('viewport','width=device-width, initial-scale=1');
}
function page_(mode){ const t=HtmlService.createTemplateFromFile('App'); t.mode=mode; return t.evaluate(); }

function donneesApp(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), list=ss.getSheetByName(LAZERISTES);
  const laz=list?list.getRange('A5:A1000').getDisplayValues().flat().map(s=>s.trim()).filter(Boolean):[];
  return {aujourdhui:aujourdhui_(),patients:lirePatients_(),lazeristes:[...new Set(laz)],url:ss.getUrl(),gid:ss.getSheetByName(REGISTRE).getSheetId()};
}
function ajouterPatient(p){
  return avecVerrou_(()=>{
    const v=preparer_(p), a=aujourdhui_();
    const ligne=ajouterLigne_(SpreadsheetApp.getActiveSpreadsheet(),[dateSerie_(a[0]+'-'+a[1]+'-'+a[2])].concat(v));
    return {ligne:ligne,data:donneesApp()};
  });
}
function modifierPatient(ligne,emp,p){
  return avecVerrou_(()=>{
    const sh=verifierLigne_(ligne,emp), v=preparer_(p);
    formaterLigne_(sh,ligne);
    sh.getRange(ligne,3,1,12).setValues([v]);
    sh.getRange(ligne,1).setValue(p.termine===true);
    return donneesApp();
  });
}
function marquerTermine(ligne,emp,termine){
  return avecVerrou_(()=>{ verifierLigne_(ligne,emp).getRange(ligne,1).setValue(termine===true); return donneesApp(); });
}
function supprimerPatient(ligne,emp){
  return avecVerrou_(()=>{ verifierLigne_(ligne,emp).deleteRow(ligne); return donneesApp(); });
}
function allerALaLigne(ligne){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sh=ss.getSheetByName(REGISTRE);
  ss.setActiveSheet(sh); sh.getRange(ligne,1,1,14).activate();
}

/* ───────────── Écriture dans le registre ───────────── */
// Écrit une nouvelle ligne (valeurs B..N) à la première ligne vide du registre et renvoie son numéro.
function ajouterLigne_(ss,valeurs){
  const sh=ss.getSheetByName(REGISTRE), max=sh.getMaxRows();
  const noms=sh.getRange(2,3,max-1,1).getDisplayValues(); const idx=noms.findIndex(x=>!String(x[0]).trim());
  const row=idx<0?max+1:idx+2;
  if(row>max) etendreRegistre_(sh,50);
  sh.getRange(row,1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()).setValue(false);
  formaterLigne_(sh,row);
  sh.getRange(row,2,1,13).setValues([valeurs]);
  return row;
}
function formaterLigne_(sh,row){ sh.getRange(row,2,1,13).setNumberFormats([FORMATS]); }
// Registre plein : ajoute nb lignes avec le même thème, puis étend les couleurs conditionnelles et le filtre.
function etendreRegistre_(sh,nb){
  const debut=sh.getMaxRows()+1;
  sh.insertRowsAfter(sh.getMaxRows(),nb);
  sh.getRange(debut-1,1,1,14).copyTo(sh.getRange(debut,1,nb,14),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
  sh.getRange(debut,1,nb,1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sh.setRowHeights(debut,nb,28);
  reglesRegistre_(sh);
  // Un filtre Sheets ne s'agrandit pas tout seul ; le recréer réinitialise seulement les critères affichés.
  const f=sh.getFilter();
  if(f){ f.remove(); sh.getRange(1,1,sh.getMaxRows(),14).createFilter(); }
}
function avecVerrou_(fn){
  const lock=LockService.getDocumentLock(); lock.waitLock(20000);
  try{ return fn(); } finally{ lock.releaseLock(); }
}
// Empreinte = date d'enregistrement + nom + prénom : protège contre une ligne déplacée ou supprimée entre-temps.
function empreinte_(t){ return [t[1],t[2],t[3]].map(s=>String(s).trim()).join('|'); }
function verifierLigne_(ligne,emp){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTRE);
  ligne=Number(ligne);
  if(!(ligne>=2&&ligne<=sh.getLastRow())||empreinte_(sh.getRange(ligne,1,1,4).getDisplayValues()[0])!==emp)
    throw new Error('Cette fiche a changé dans le registre entre-temps. Rechargez les données puis réessayez.');
  return sh;
}
// Valide la saisie du formulaire et renvoie les valeurs C..N prêtes pour le registre.
function preparer_(p){
  const txt=x=>String(x==null?'':x).trim(), nom=txt(p.nom), prenom=txt(p.prenom);
  if(!nom||!prenom) throw new Error('Le nom et le prénom sont obligatoires.');
  const num=(x,lib,entier)=>{
    if(txt(x)==='') return '';
    const n=Number(txt(x).replace(',','.'));
    if(isNaN(n)||n<0||(entier&&n%1)) throw new Error(lib+' invalide.');
    return n;
  };
  return [nom,prenom,txt(p.region),txt(p.tel),heureSerie_(p.arrivee),dateSerie_(p.date),heureSerie_(p.heure),txt(p.laz),txt(p.params),num(p.seances,'Nombre de séances',true),num(p.prix,'Prix'),txt(p.remarques)];
}
// Numéros de série Sheets (jours depuis le 30/12/1899) : indépendants du fuseau horaire et de la langue du tableur.
function dateSerie_(s){
  if(!s) return '';
  const m=String(s).match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); if(!m) throw new Error('Date invalide.');
  return (Date.UTC(+m[1],m[2]-1,+m[3])-Date.UTC(1899,11,30))/864e5;
}
function heureSerie_(s){
  if(!s) return '';
  const m=String(s).match(/^(\d{1,2}):(\d{2})$/); if(!m||+m[1]>23||+m[2]>59) throw new Error('Heure invalide.');
  return (+m[1]*60+ +m[2])/1440;
}
