const LTJ_SPREADSHEET_ID = '1gRv9T8sq3aJUA30Y9GnqWsMBIr1_C84afXoyCZXjh8w';
const ONESIGNAL_APP_ID = 'e7edffbe-6a42-4019-ae38-9879feca5b46';
const PORTAL_URL = 'https://ltjintermediacao.github.io/ltjintermediacao/#calendar-section';

function enviarEventosDeHoje() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ONESIGNAL_API_KEY');
  if (!apiKey) throw new Error('Configure ONESIGNAL_API_KEY nas Propriedades do script.');
  const sh = SpreadsheetApp.openById(LTJ_SPREADSHEET_ID).getSheetByName('CALENDARIO');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(v => String(v).trim().toLowerCase());
  const col = nome => headers.indexOf(nome.toLowerCase());
  const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  for (let i=1;i<values.length;i++) {
    const row=values[i], data=row[col('Data')], evento=String(row[col('Evento')]||'').trim();
    const horario=String(row[col('Horário')]||'').trim(), local=String(row[col('Local')]||'').trim();
    const ativo=String(row[col('Ativo')]||'SIM').trim().toUpperCase();
    const notificar=String(row[col('Notificar no dia')]||'SIM').trim().toUpperCase();
    if(!data||!evento||ativo==='NÃO'||notificar==='NÃO') continue;
    const dataEvento=Utilities.formatDate(new Date(data),'America/Sao_Paulo','yyyy-MM-dd');
    if(dataEvento!==hoje) continue;
    const chave='evento_enviado_'+i+'_'+hoje, props=PropertiesService.getScriptProperties();
    if(props.getProperty(chave)) continue;
    const detalhes=[horario,local].filter(Boolean).join(' • ');
    enviarPushOneSignal('📅 Evento de hoje: '+evento, detalhes||'Confira o calendário do Portal LTJ.');
    props.setProperty(chave,'SIM');
  }
}

function enviarPushOneSignal(titulo,mensagem) {
  const apiKey=PropertiesService.getScriptProperties().getProperty('ONESIGNAL_API_KEY');
  const payload={app_id:ONESIGNAL_APP_ID,target_channel:'push',included_segments:['Subscribed Users'],
    headings:{en:titulo,pt:titulo},contents:{en:mensagem,pt:mensagem},url:PORTAL_URL};
  const response=UrlFetchApp.fetch('https://api.onesignal.com/notifications',{
    method:'post',contentType:'application/json',
    headers:{Authorization:'Key '+apiKey},payload:JSON.stringify(payload),muteHttpExceptions:true});
  if(response.getResponseCode()>=300) throw new Error('OneSignal: '+response.getContentText());
}

function criarGatilhoDiario() {
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='enviarEventosDeHoje')
    .forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarEventosDeHoje').timeBased().everyDays(1).atHour(8).create();
}
