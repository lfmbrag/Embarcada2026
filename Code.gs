var SPREADSHEET_ID = '1jPOKTCVh7LCFRYUgdiwHHIPYwg5aW0kaZ4Jq7poWiFY';

function getPlanilha_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Roteador do Web App: decide qual página HTML servir de acordo com
 * o parâmetro ?page= da URL. Página padrão: "inicio".
 */
function doGet(e) {
  var pagina = (e && e.parameter && e.parameter.page) || 'inicio';

  var mapaDeArquivos = {
    'inicio': 'Inicio',
    'triagem': 'FormularioTriagem',
    'exames': 'FormularioExames',
    'anamnese': 'FormularioAnamnese',
    'fisioterapia': 'FormularioFisioterapia',
    'mamografia': 'FormularioMamografia',
    'consulta': 'ConsultaAtendimento'
  };

  var arquivo = mapaDeArquivos[pagina] || 'Inicio';

  var template = HtmlService.createTemplateFromFile(arquivo);
  // baseUrl = URL pública real do Web App (ex: https://script.google.com/macros/s/SEU_ID/exec)
  // É repassada ao menu de navegação para montar links absolutos e confiáveis,
  // em vez de links relativos que quebram dentro do iframe interno do Apps Script.
  template.baseUrl = ScriptApp.getService().getUrl();

  return template
    .evaluate()
    .setTitle('Atendimento Comunitário')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // ALLOWALL é necessário para o Google Sites conseguir exibir esta
    // página dentro de um <iframe> ao incorporar por URL.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Permite incluir arquivos HTML separados (Nav, Stylesheet, JavaScript
 * de cada página) dentro dos templates, via <?!= include('Nome'); ?>
 * Aceita um segundo parâmetro opcional com variáveis a repassar para o
 * template incluído, ex: <?!= include('Nav', {baseUrl: baseUrl}); ?>
 *
 * Se o arquivo não existir (nome digitado errado, arquivo não criado no
 * projeto, etc.), em vez de derrubar a página inteira com um erro genérico,
 * mostra um aviso visível dizendo exatamente qual arquivo está faltando —
 * assim fica fácil identificar o problema sem precisar abrir o log de
 * execuções.
 */
function include(filename, dados) {
  try {
    var tmpl = HtmlService.createTemplateFromFile(filename);
    if (dados) {
      Object.keys(dados).forEach(function (chave) {
        tmpl[chave] = dados[chave];
      });
    }
    return tmpl.evaluate().getContent();
  } catch (err) {
    return '<div style="margin:16px 0;padding:14px 18px;border:1px solid #b3452f;' +
      'border-radius:8px;background:#fdf2ef;color:#b3452f;font-family:sans-serif;font-size:13px;">' +
      '<strong>Arquivo "' + filename + '" não encontrado ou com erro.</strong><br>' +
      'Verifique se ele existe no projeto do Apps Script com esse nome exato ' +
      '(sem espaços, sem ".html", maiúsculas/minúsculas idênticas).<br>' +
      '<span style="opacity:0.75;">Detalhe técnico: ' + err.message + '</span>' +
      '</div>';
  }
}

/**
 * (Opcional) Cria um menu dentro do próprio Google Sheets para abrir
 * o Web App rapidamente, caso o script esteja vinculado à planilha.
 * Não é necessário para o funcionamento do Web App/Google Sites.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Atendimento Comunitário')
    .addItem('Abrir Web App', 'abrirWebAppNoNavegador_')
    .addToUi();
}

function abrirWebAppNoNavegador_() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Abrindo...');
}

/* =====================================================================
 * UTILITÁRIOS COMUNS
 * ===================================================================== */

function getOrCreateSheet_(nomeAba, cabecalho) {
  var ss = getPlanilha_();
  var sheet = ss.getSheetByName(nomeAba);

  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(cabecalho);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, cabecalho.length)
      .setFontWeight('bold')
      .setBackground('#0f3d3e')
      .setFontColor('#ffffff');
    sheet.autoResizeColumns(1, cabecalho.length);
  }

  return sheet;
}

function combinarComOutro_(valor, valorOutro) {
  if (!valor) return '';
  if (valor.indexOf('Outro') !== -1 && valorOutro) {
    return valor + ' (' + valorOutro + ')';
  }
  return valor;
}

/* =====================================================================
 * TRIAGEM  ->  aba "Registros"
 * ===================================================================== */

var NOME_ABA_TRIAGEM = 'Registros';

var CABECALHO_TRIAGEM = [
  'Data/Hora do registro', 'Comunidade', 'Tipo de atendimento', 'Tipo de atendimento procurado',
  'Número do atendimento', 'Nome do paciente', 'Nome da mãe', 'Número do documento',
  'Número do cartão SUS', 'Idade', 'Data de nascimento', 'Número de telefone',
  'Autodeclaração étnica', 'Estado civil', 'Sexo', 'Gestante', 'Idade gestacional',
  'Puérpera', 'Tempo desde o parto', 'Possui filhos', 'Quantos filhos', 'Idade dos filhos',
  'Carteira de vacinação apresentada', 'Carteira completa', 'Vacinas realizadas',
  'Alteração ocular visível', 'Estudou', 'Escolaridade', 'Trabalho/Ofício',
  'Internações anteriores', 'Motivo da internação', 'Histórico familiar',
  'Água utilizada vem de', 'Água recebe tratamento', 'Como é o tratamento da água',
  'Possui banheiro/sanitário', 'Destino do lixo', 'Tabagismo',
  'Forma de exposição ao tabagismo', 'Formas de consumo de tabaco',
  'Carga diária (tabagismo)', 'Tempo de tabagismo', 'Tempo como ex-tabagista',
  'Consumo de álcool', 'Frequência do consumo de álcool', 'Uso de outra substância',
  'Qual substância', 'Peso (kg)', 'Altura (cm)', 'Temperatura (°C)',
  'Frequência Cardíaca (bpm)', 'Pressão Arterial (mmHg)', 'Saturação de Oxigênio (%)',
  'Circunferência abdominal (cm)', 'Outras observações', 'Paciente identificado (confirmado)'
];

function registrarAtendimento(dados) {
  try {
    var sheet = getOrCreateSheet_(NOME_ABA_TRIAGEM, CABECALHO_TRIAGEM);
    sheet.appendRow([
      new Date(),
      combinarComOutro_(dados.comunidade || '', dados.comunidadeOutro),
      dados.tipoAtendimento || '',
      combinarComOutro_(dados.tipoAtendimentoProcura || '', dados.tipoAtendimentoProcuraOutro),
      dados.numeroAtendimento || '', dados.nome || '', dados.nomeMae || '',
      dados.documento || '', dados.cartaoSus || '', dados.idade || '', dados.dataNascimento || '',
      dados.telefone || '',
      combinarComOutro_(dados.etnia || '', dados.etniaOutro),
      combinarComOutro_(dados.estadoCivil || '', dados.estadoCivilOutro),
      combinarComOutro_(dados.sexo || '', dados.sexoOutro),
      combinarComOutro_(dados.gestante || '', dados.gestanteOutro),
      dados.idadeGestacional || '',
      combinarComOutro_(dados.puerpera || '', dados.puerperaOutro),
      dados.tempoParto || '',
      dados.possuiFilhos || '',
      combinarComOutro_(dados.quantosFilhos || '', dados.quantosFilhosOutro),
      dados.idadeFilhos || '',
      combinarComOutro_(dados.carteiraVacinacao || '', dados.carteiraVacinacaoOutro),
      combinarComOutro_(dados.carteiraCompleta || '', dados.carteiraCompletaOutro),
      combinarComOutro_(dados.vacinasRealizadas || '', dados.vacinasRealizadasOutro),
      combinarComOutro_(dados.alteracaoOcular || '', dados.alteracaoOcularOutro),
      combinarComOutro_(dados.estudou || '', dados.estudouOutro),
      dados.escolaridade || '',
      combinarComOutro_(dados.trabalho || '', dados.trabalhoOutro),
      combinarComOutro_(dados.internacoes || '', dados.internacoesOutro),
      dados.motivoInternacao || '',
      combinarComOutro_(dados.historicoFamiliar || '', dados.historicoFamiliarOutro),
      combinarComOutro_(dados.fonteAgua || '', dados.fonteAguaOutro),
      combinarComOutro_(dados.tratamentoAgua || '', dados.tratamentoAguaOutro),
      combinarComOutro_(dados.comoTratamento || '', dados.comoTratamentoOutro),
      combinarComOutro_(dados.banheiro || '', dados.banheiroOutro),
      combinarComOutro_(dados.destinoLixo || '', dados.destinoLixoOutro),
      combinarComOutro_(dados.tabagismo || '', dados.tabagismoOutro),
      dados.formaExposicao || '',
      combinarComOutro_(dados.formasConsumo || '', dados.formasConsumoOutro),
      dados.cargaDiaria || '', dados.tempoTabagismo || '', dados.tempoExTabagista || '',
      dados.alcool || '',
      combinarComOutro_(dados.frequenciaAlcool || '', dados.frequenciaAlcoolOutro),
      combinarComOutro_(dados.outraSubstancia || '', dados.outraSubstanciaOutro),
      dados.qualSubstancia || '',
      dados.peso || '', dados.altura || '', dados.temperatura || '',
      dados.frequenciaCardiaca || '', dados.pressaoArterial || '', dados.saturacaoOxigenio || '',
      dados.circunferenciaAbdominal || '', dados.observacoes || '',
      dados.identificacaoConfirmada ? 'Sim' : 'Não'
    ]);
    return { sucesso: true };
  } catch (err) {
    return { sucesso: false, erro: err.message };
  }
}

/* =====================================================================
 * EXAMES  ->  aba "Exames"
 * ===================================================================== */

var NOME_ABA_EXAMES = 'Exames';

var CABECALHO_EXAMES = [
  'Data/Hora do registro', 'Tipo de atendimento', 'Número do atendimento', 'Nome',
  'Médico que solicitou', 'Exames solicitados', 'Material de análise', 'HIV',
  'Sífilis', 'Hepatite B (HBsAG)', 'Hepatite C (Anti-HCV)',
  'Hemácias (milhões/µL)', 'Hemoglobina (g/dL)', 'Hematócrito (%)', 'VCM (fL)',
  'HCM (pg)', 'CHCM (g/dL)', 'RDW (%)', 'Morfologia das hemácias',
  'Leucócitos totais', 'Neutrófilos (%)', 'Linfócitos (%)', 'Monócitos (%)',
  'Eosinófilos (%)', 'Basófilos (%)', 'Plaquetas', 'Aspecto (EQU)',
  'Cor (EQU)', 'pH', 'Densidade', 'Proteínas', 'Glicose', 'Corpos cetônicos',
  'Bilirrubina', 'Leucócitos (EQU)', 'Nitrito', 'Sangue/Hemoglobina (EQU)',
  'Urobilinogênio'
];

function registrarExame(dados) {
  try {
    var sheet = getOrCreateSheet_(NOME_ABA_EXAMES, CABECALHO_EXAMES);

    var morfologia = dados.morfologiaHemacias === 'Outro' && dados.morfologiaHemaciasOutro
      ? 'Outro: ' + dados.morfologiaHemaciasOutro
      : (dados.morfologiaHemacias || '');

    var cor = dados.corEqu === 'Outro' && dados.corEquOutro
      ? 'Outro: ' + dados.corEquOutro
      : (dados.corEqu || '');

    sheet.appendRow([
      new Date(),
      dados.tipoAtendimento || '', dados.numeroAtendimento || '', dados.nome || '',
      dados.medico || '', dados.examesSolicitados || '',
      combinarComOutro_(dados.materialAnalise || '', dados.materialAnaliseOutro),
      dados.hiv || '', dados.sifilis || '', dados.hepatiteB || '', dados.hepatiteC || '',
      dados.hemacias || '', dados.hemoglobina || '', dados.hematocrito || '',
      dados.vcm || '', dados.hcm || '', dados.chcm || '', dados.rdw || '', morfologia,
      dados.leucocitosTotais || '', dados.neutrofilos || '', dados.linfocitos || '',
      dados.monocitos || '', dados.eosinofilos || '', dados.basofilos || '',
      dados.plaquetas || '', dados.aspectoEqu || '', cor, dados.ph || '',
      dados.densidade || '', dados.proteinas || '', dados.glicoseEqu || '',
      dados.corposCetonicos || '', dados.bilirrubina || '', dados.leucocitosEqu || '',
      dados.nitrito || '', dados.sangueHemoglobinaEqu || '', dados.urobilinogenio || ''
    ]);
    return { sucesso: true };
  } catch (err) {
    return { sucesso: false, erro: err.message };
  }
}

/* =====================================================================
 * ANAMNESE / CONSULTA  ->  aba "Anamnese"
 * ===================================================================== */

var NOME_ABA_ANAMNESE = 'Anamnese';

var CABECALHO_ANAMNESE = [
  'Data/Hora do registro', 'Tipo de atendimento', 'Número do atendimento', 'Nome',
  'Médico responsável', 'Possui alguma doença?', 'Têm ou teve câncer?', 'Câncer - tipo/onde',
  'Câncer - quimioterapia', 'Câncer - radioterapia', 'Câncer - observações',
  'Já fez cirurgia?', 'Qual cirurgia', 'Utiliza medicação contínua?',
  'Qual medicamento de uso contínuo', 'Tem alergia a medicamento?', 'Qual alergia',
  'Histórico últimos 3 meses',
  'Estado geral', 'Cabeça e pescoço', 'Cabeça e pescoço - especifique',
  'Cardiovascular', 'Cardiovascular - especifique', 'Abdômen', 'Abdômen - especifique',
  'Pele', 'Pele - especifique', 'Membros', 'Membros - especifique',
  'Neurológico', 'Neurológico - especifique', 'Olhos', 'Olhos - especifique',
  'Peso ao nascer', 'Doenças prévias (infantil)', 'Prematuridade',
  'Semanas gestacionais', 'Internação neonatal', 'Internação neonatal - especifique',
  'Intercorrências na infância', 'Quais intercorrências',
  'Breve descrição do caso', 'Impressão diagnóstica', 'Observações', 'Conduta'
];

function registrarAnamnese(dados) {
  try {
    var sheet = getOrCreateSheet_(NOME_ABA_ANAMNESE, CABECALHO_ANAMNESE);
    sheet.appendRow([
      new Date(),
      dados.tipoAtendimento || '', dados.numeroAtendimento || '', dados.nome || '',
      dados.medico || '',
      combinarComOutro_(dados.doencas || '', dados.doencasOutro),
      dados.cancer || '', dados.cancerTipo || '', dados.quimioterapia || '',
      dados.radioterapia || '', dados.cancerObservacoes || '',
      dados.cirurgia || '', dados.qualCirurgia || '',
      dados.medicacaoContinua || '', dados.medicamentoContinuo || '',
      combinarComOutro_(dados.alergiaMedicamento || '', dados.alergiaMedicamentoOutro), dados.qualAlergia || '',
      combinarComOutro_(dados.historico3meses || '', dados.historico3mesesOutro),
      dados.estadoGeral || '',
      dados.cabecaPescoco || '', dados.cabecaPescocoEspecifique || '',
      dados.cardiovascular || '', dados.cardiovascularEspecifique || '',
      dados.abdomen || '', dados.abdomenEspecifique || '',
      dados.pele || '', dados.peleEspecifique || '',
      dados.membros || '', dados.membrosEspecifique || '',
      dados.neurologico || '', dados.neurologicoEspecifique || '',
      dados.olhos || '', dados.olhosEspecifique || '',
      dados.pesoNascer || '',
      combinarComOutro_(dados.doencasPreviasInfantil || '', dados.doencasPreviasInfantilOutro),
      dados.prematuridade || '', dados.semanasGestacionais || '',
      dados.internacaoNeonatal || '', dados.internacaoNeonatalEspecifique || '',
      combinarComOutro_(dados.intercorrenciasInfancia || '', dados.intercorrenciasInfanciaOutro),
      dados.quaisIntercorrencias || '',
      dados.descricaoCaso || '', dados.impressaoDiagnostica || '', dados.observacoesConsulta || '',
      combinarComOutro_(dados.conduta || '', dados.condutaOutro)
    ]);
    return { sucesso: true };
  } catch (err) {
    return { sucesso: false, erro: err.message };
  }
}

/* =====================================================================
 * FISIOTERAPIA  ->  aba "Fisioterapia"
 * ===================================================================== */

var NOME_ABA_FISIO = 'Fisioterapia';

var CABECALHO_FISIO = [
  'Data/Hora do registro', 'Tipo de atendimento', 'Número do atendimento', 'Nome',
  'Fisioterapeuta', 'O trabalho exige', 'Principal meio de transporte',
  'AVC', 'Hipertensão', 'Diabetes', 'Fratura', 'Cirurgias', 'Gestante',
  'Medicamentos contínuos', 'Motivo do atendimento', 'Início dos sintomas',
  'Como começaram os sintomas', 'Piora com', 'Melhora com',
  'Escala da dor (EVN 0-10)', 'Características da dor', 'Local da dor',
  'A dor impede', 'Postura', 'Marcha', 'Amplitude - Coluna cervical',
  'Amplitude - Ombro', 'Amplitude - Cotovelo', 'Amplitude - Punho/mão',
  'Amplitude - Coluna Torácica', 'Amplitude - Coluna Lombar',
  'Amplitude - Quadril', 'Amplitude - Joelho', 'Amplitude - Tornozelo/Pé',
  'Força - Coluna cervical', 'Força - Ombro', 'Força - Cotovelo',
  'Força - Punho/mão', 'Força - Coluna Torácica', 'Força - Coluna Lombar',
  'Força - Joelho', 'Força - Tornozelo/Pé', 'Sensibilidade', 'Equilíbrio',
  'Edema', 'Local do edema', 'Diagnóstico Cinético-funcional',
  'Hipótese fisioterapêutica', 'Conduta realizada', 'Encaminhamento',
  'Orientações domiciliares', 'Observações'
];

function registrarFisioterapia(dados) {
  try {
    var sheet = getOrCreateSheet_(NOME_ABA_FISIO, CABECALHO_FISIO);

    sheet.appendRow([
      new Date(),
      dados.tipoAtendimento || '', dados.numeroAtendimento || '', dados.nome || '',
      dados.fisioterapeuta || '',
      combinarComOutro_(dados.trabalhoExige || '', dados.trabalhoExigeOutro),
      combinarComOutro_(dados.transporte || '', dados.transporteOutro),
      dados.avc || '', dados.hipertensao || '', dados.diabetes || '', dados.fratura || '',
      dados.cirurgias || '', dados.gestante || '', dados.medicamentosContinuos || '',
      combinarComOutro_(dados.motivoAtendimento || '', dados.motivoAtendimentoOutro),
      dados.inicioSintomas || '', dados.comoComecou || '',
      combinarComOutro_(dados.pioraCom || '', dados.pioraComOutro),
      combinarComOutro_(dados.melhoraCom || '', dados.melhoraComOutro),
      dados.escalaDor || '',
      combinarComOutro_(dados.caracteristicasDor || '', dados.caracteristicasDorOutro),
      combinarComOutro_(dados.localDor || '', dados.localDorOutro),
      combinarComOutro_(dados.dorImpede || '', dados.dorImpedeOutro),
      combinarComOutro_(dados.postura || '', dados.posturaOutro),
      combinarComOutro_(dados.marcha || '', dados.marchaOutro),
      dados.amplitudeCervical || '', dados.amplitudeOmbro || '', dados.amplitudeCotovelo || '',
      dados.amplitudePunhoMao || '', dados.amplitudeToracica || '', dados.amplitudeLombar || '',
      dados.amplitudeQuadril || '', dados.amplitudeJoelho || '', dados.amplitudeTornozeloPe || '',
      dados.forcaCervical || '', dados.forcaOmbro || '', dados.forcaCotovelo || '',
      dados.forcaPunhoMao || '', dados.forcaToracica || '', dados.forcaLombar || '',
      dados.forcaJoelho || '', dados.forcaTornozeloPe || '', dados.sensibilidade || '',
      dados.equilibrio || '', dados.edema || '', dados.localEdema || '',
      dados.diagnosticoCinetico || '', dados.hipoteseFisioterapeutica || '',
      combinarComOutro_(dados.condutaRealizada || '', dados.condutaRealizadaOutro),
      combinarComOutro_(dados.encaminhamento || '', dados.encaminhamentoOutro),
      combinarComOutro_(dados.orientacoesDomiciliares || '', dados.orientacoesDomiciliaresOutro),
      dados.observacoes || ''
    ]);
    return { sucesso: true };
  } catch (err) {
    return { sucesso: false, erro: err.message };
  }
}

/* =====================================================================
 * MAMOGRAFIA  ->  aba "Mamografia"
 * ===================================================================== */

var NOME_ABA_MAMOGRAFIA = 'Mamografia';

var CABECALHO_MAMOGRAFIA = [
  'Data/Hora do registro', 'Tipo de atendimento', 'Número do atendimento', 'Nome',
  'Tipo de mamografia', 'Encaminhamento ao exame', 'Motivo do encaminhamento',
  'Mamografia anterior', 'Data da última mamografia (estimativa)', 'Resultado anterior',
  'Cirurgia mamária prévia', 'Radioterapia prévia', 'Implante mamário',
  'História pessoal de câncer de mama', 'História familiar de câncer de mama',
  'Grau de parentesco', 'Terapia hormonal', 'Utiliza anticoncepcional oral', 'Lactação',
  'Dados clínicos', 'Paciente sintomática?',
  'Achados mamográficos - Mama direita', 'Achados mamográficos - Mama esquerda',
  'Tipo de nódulo', 'Tamanho do nódulo', 'Densidade mamária (BI-RADS)',
  'Classificação BI-RADS', 'Conclusão do exame', 'Recomendação', 'Observações adicionais'
];

function registrarMamografia(dados) {
  try {
    var sheet = getOrCreateSheet_(NOME_ABA_MAMOGRAFIA, CABECALHO_MAMOGRAFIA);
    sheet.appendRow([
      new Date(),
      dados.tipoAtendimento || '', dados.numeroAtendimento || '', dados.nome || '',
      dados.tipoMamografia || '',
      combinarComOutro_(dados.encaminhamento || '', dados.encaminhamentoOutro),
      dados.motivoEncaminhamento || '',
      dados.mamografiaAnterior || '', dados.dataUltimaMamografia || '', dados.resultadoAnterior || '',
      combinarComOutro_(dados.cirurgiaMamaria || '', dados.cirurgiaMamariaOutro),
      combinarComOutro_(dados.radioterapiaPrevia || '', dados.radioterapiaPreviaOutro),
      combinarComOutro_(dados.implanteMamario || '', dados.implanteMamarioOutro),
      combinarComOutro_(dados.historiaPessoalCancer || '', dados.historiaPessoalCancerOutro),
      combinarComOutro_(dados.historiaFamiliarCancer || '', dados.historiaFamiliarCancerOutro),
      combinarComOutro_(dados.grauParentesco || '', dados.grauParentescoOutro),
      dados.terapiaHormonal || '', dados.anticoncepcional || '', dados.lactacao || '',
      combinarComOutro_(dados.dadosClinicos || '', dados.dadosClinicosOutro),
      dados.sintomatica || '',
      combinarComOutro_(dados.achadosDireita || '', dados.achadosDireitaOutro),
      combinarComOutro_(dados.achadosEsquerda || '', dados.achadosEsquerdaOutro),
      dados.noduloTipo || '', dados.noduloTamanho || '',
      dados.densidadeMamaria || '', dados.biRads || '',
      dados.conclusaoExame || '',
      combinarComOutro_(dados.recomendacao || '', dados.recomendacaoOutro),
      dados.observacoesAdicionais || ''
    ]);
    return { sucesso: true };
  } catch (err) {
    return { sucesso: false, erro: err.message };
  }
}

/* =====================================================================
 * BUSCAR PACIENTE  ->  lê de uma ou mais abas, filtrando por "Número"
 * ===================================================================== */

// Configuração de cada categoria buscável: aba, grupos de exibição e
// o nome da coluna que identifica o paciente (para o título do cartão).
var CATEGORIAS_BUSCA = {
  triagem: {
    rotulo: 'Triagem',
    aba: NOME_ABA_TRIAGEM,
    campoNumero: 'Número do atendimento',
    campoTitulo: 'Nome do paciente',
    grupos: [
      { titulo: 'Identificação do atendimento', campos: ['Data/Hora do registro', 'Comunidade', 'Tipo de atendimento', 'Tipo de atendimento procurado', 'Número do atendimento'] },
      { titulo: 'Dados do paciente', campos: ['Nome do paciente', 'Nome da mãe', 'Número do documento', 'Número do cartão SUS', 'Idade', 'Data de nascimento', 'Número de telefone', 'Autodeclaração étnica', 'Estado civil', 'Sexo'] },
      { titulo: 'Saúde reprodutiva', campos: ['Gestante', 'Idade gestacional', 'Puérpera', 'Tempo desde o parto', 'Possui filhos', 'Quantos filhos', 'Idade dos filhos'] },
      { titulo: 'Vacinação e antecedentes', campos: ['Carteira de vacinação apresentada', 'Carteira completa', 'Vacinas realizadas', 'Alteração ocular visível', 'Estudou', 'Escolaridade', 'Trabalho/Ofício', 'Internações anteriores', 'Motivo da internação', 'Histórico familiar'] },
      { titulo: 'Condições de moradia', campos: ['Água utilizada vem de', 'Água recebe tratamento', 'Como é o tratamento da água', 'Possui banheiro/sanitário', 'Destino do lixo'] },
      { titulo: 'Hábitos de vida', campos: ['Tabagismo', 'Forma de exposição ao tabagismo', 'Formas de consumo de tabaco', 'Carga diária (tabagismo)', 'Tempo de tabagismo', 'Tempo como ex-tabagista', 'Consumo de álcool', 'Frequência do consumo de álcool', 'Uso de outra substância', 'Qual substância'] },
      { titulo: 'Sinais vitais', campos: ['Peso (kg)', 'Altura (cm)', 'Temperatura (°C)', 'Frequência Cardíaca (bpm)', 'Pressão Arterial (mmHg)', 'Saturação de Oxigênio (%)', 'Circunferência abdominal (cm)'] },
      { titulo: 'Observações', campos: ['Outras observações', 'Paciente identificado (confirmado)'] }
    ]
  },
  exames: {
    rotulo: 'Exames',
    aba: NOME_ABA_EXAMES,
    campoNumero: 'Número do atendimento',
    campoTitulo: 'Nome',
    grupos: [
      { titulo: 'Identificação', campos: ['Data/Hora do registro', 'Tipo de atendimento', 'Número do atendimento', 'Médico que solicitou', 'Exames solicitados', 'Material de análise'] },
      { titulo: 'Testes rápidos', campos: ['HIV', 'Sífilis', 'Hepatite B (HBsAG)', 'Hepatite C (Anti-HCV)'] },
      { titulo: 'Hemograma', campos: ['Hemácias (milhões/µL)', 'Hemoglobina (g/dL)', 'Hematócrito (%)', 'VCM (fL)', 'HCM (pg)', 'CHCM (g/dL)', 'RDW (%)', 'Morfologia das hemácias', 'Leucócitos totais', 'Neutrófilos (%)', 'Linfócitos (%)', 'Monócitos (%)', 'Eosinófilos (%)', 'Basófilos (%)', 'Plaquetas'] },
      { titulo: 'EQU', campos: ['Aspecto (EQU)', 'Cor (EQU)', 'pH', 'Densidade', 'Proteínas', 'Glicose', 'Corpos cetônicos', 'Bilirrubina', 'Leucócitos (EQU)', 'Nitrito', 'Sangue/Hemoglobina (EQU)', 'Urobilinogênio'] }
    ]
  },
  anamnese: {
    rotulo: 'Consulta',
    aba: NOME_ABA_ANAMNESE,
    campoNumero: 'Número do atendimento',
    campoTitulo: 'Nome',
    grupos: [
      { titulo: 'Identificação', campos: ['Data/Hora do registro', 'Tipo de atendimento', 'Número do atendimento', 'Médico responsável'] },
      { titulo: 'História prévia', campos: ['Possui alguma doença?', 'Têm ou teve câncer?', 'Câncer - tipo/onde', 'Câncer - quimioterapia', 'Câncer - radioterapia', 'Câncer - observações', 'Já fez cirurgia?', 'Qual cirurgia', 'Utiliza medicação contínua?', 'Qual medicamento de uso contínuo', 'Tem alergia a medicamento?', 'Qual alergia', 'Histórico últimos 3 meses'] },
      { titulo: 'Exame físico', campos: ['Estado geral', 'Cabeça e pescoço', 'Cabeça e pescoço - especifique', 'Cardiovascular', 'Cardiovascular - especifique', 'Abdômen', 'Abdômen - especifique', 'Pele', 'Pele - especifique', 'Membros', 'Membros - especifique', 'Neurológico', 'Neurológico - especifique', 'Olhos', 'Olhos - especifique'] },
      { titulo: 'Adicionais de atendimento infantil', campos: ['Peso ao nascer', 'Doenças prévias (infantil)', 'Prematuridade', 'Semanas gestacionais', 'Internação neonatal', 'Internação neonatal - especifique', 'Intercorrências na infância', 'Quais intercorrências'] },
      { titulo: 'Consulta médica', campos: ['Breve descrição do caso', 'Impressão diagnóstica', 'Observações', 'Conduta'] }
    ]
  },
  fisioterapia: {
    rotulo: 'Fisioterapia',
    aba: NOME_ABA_FISIO,
    campoNumero: 'Número do atendimento',
    campoTitulo: 'Nome',
    grupos: [
      { titulo: 'Identificação', campos: ['Data/Hora do registro', 'Tipo de atendimento', 'Número do atendimento', 'Fisioterapeuta', 'O trabalho exige', 'Principal meio de transporte'] },
      { titulo: 'Histórico da doença pregressa', campos: ['AVC', 'Hipertensão', 'Diabetes', 'Fratura', 'Cirurgias', 'Gestante', 'Medicamentos contínuos'] },
      { titulo: 'Histórico da doença atual', campos: ['Motivo do atendimento', 'Início dos sintomas', 'Como começaram os sintomas', 'Piora com', 'Melhora com'] },
      { titulo: 'Avaliação da dor', campos: ['Escala da dor (EVN 0-10)', 'Características da dor', 'Local da dor', 'A dor impede'] },
      { titulo: 'Avaliação da especialidade', campos: ['Postura', 'Marcha', 'Amplitude - Coluna cervical', 'Amplitude - Ombro', 'Amplitude - Cotovelo', 'Amplitude - Punho/mão', 'Amplitude - Coluna Torácica', 'Amplitude - Coluna Lombar', 'Amplitude - Quadril', 'Amplitude - Joelho', 'Amplitude - Tornozelo/Pé', 'Força - Coluna cervical', 'Força - Ombro', 'Força - Cotovelo', 'Força - Punho/mão', 'Força - Coluna Torácica', 'Força - Coluna Lombar', 'Força - Joelho', 'Força - Tornozelo/Pé', 'Sensibilidade', 'Equilíbrio', 'Edema', 'Local do edema'] },
      { titulo: 'Conclusão de atendimento', campos: ['Diagnóstico Cinético-funcional', 'Hipótese fisioterapêutica', 'Conduta realizada', 'Encaminhamento', 'Orientações domiciliares', 'Observações'] }
    ]
  },
  mamografia: {
    rotulo: 'Mamografia',
    aba: NOME_ABA_MAMOGRAFIA,
    campoNumero: 'Número do atendimento',
    campoTitulo: 'Nome',
    grupos: [
      { titulo: 'Identificação', campos: ['Data/Hora do registro', 'Tipo de atendimento', 'Número do atendimento', 'Tipo de mamografia', 'Encaminhamento ao exame', 'Motivo do encaminhamento'] },
      { titulo: 'Histórico da paciente', campos: ['Mamografia anterior', 'Data da última mamografia (estimativa)', 'Resultado anterior', 'Cirurgia mamária prévia', 'Radioterapia prévia', 'Implante mamário', 'História pessoal de câncer de mama', 'História familiar de câncer de mama', 'Grau de parentesco', 'Terapia hormonal', 'Utiliza anticoncepcional oral', 'Lactação'] },
      { titulo: 'Atendimento/Exame', campos: ['Dados clínicos', 'Paciente sintomática?', 'Achados mamográficos - Mama direita', 'Achados mamográficos - Mama esquerda', 'Tipo de nódulo', 'Tamanho do nódulo', 'Densidade mamária (BI-RADS)', 'Classificação BI-RADS', 'Conclusão do exame', 'Recomendação', 'Observações adicionais'] }
    ]
  }
};

/**
 * Busca os registros de um paciente (por número de atendimento) em uma
 * ou mais categorias selecionadas pelo profissional na tela de busca.
 *
 * @param {string} numero    Número do atendimento a buscar
 * @param {string[]} categorias  Subconjunto de ['triagem','exames','anamnese','fisioterapia']
 */
function buscarDadosPaciente(numero, categorias) {
  try {
    var numeroBusca = String(numero || '').trim();
    if (!numeroBusca) {
      return { sucesso: false, erro: 'Informe um número de atendimento para buscar.' };
    }
    if (!categorias || categorias.length === 0) {
      return { sucesso: false, erro: 'Selecione ao menos uma categoria para buscar (Triagem, Exames, Consulta ou Fisioterapia).' };
    }

    var ss = getPlanilha_();
    var resultadoPorCategoria = [];

    categorias.forEach(function (chave) {
      var config = CATEGORIAS_BUSCA[chave];
      if (!config) return;

      var sheet = ss.getSheetByName(config.aba);
      if (!sheet) {
        resultadoPorCategoria.push({
          categoria: chave,
          rotulo: config.rotulo,
          grupos: config.grupos,
          campoTitulo: config.campoTitulo,
          registros: [],
          avisoAbaInexistente: true
        });
        return;
      }

      var dadosPlanilha = sheet.getDataRange().getValues();
      var registros = [];

      if (dadosPlanilha.length >= 2) {
        var cabecalho = dadosPlanilha[0];
        var indiceNumero = cabecalho.indexOf(config.campoNumero);

        if (indiceNumero !== -1) {
          for (var i = 1; i < dadosPlanilha.length; i++) {
            var valorLinha = String(dadosPlanilha[i][indiceNumero]).trim();
            if (valorLinha === numeroBusca) {
              var registro = {};
              for (var j = 0; j < cabecalho.length; j++) {
                var valor = dadosPlanilha[i][j];
                if (Object.prototype.toString.call(valor) === '[object Date]') {
                  valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
                }
                registro[cabecalho[j]] = valor;
              }
              registros.push(registro);
            }
          }
        }
      }

      resultadoPorCategoria.push({
        categoria: chave,
        rotulo: config.rotulo,
        grupos: config.grupos,
        campoTitulo: config.campoTitulo,
        registros: registros
      });
    });

    return { sucesso: true, resultados: resultadoPorCategoria };
  } catch (err) {
    return { sucesso: false, erro: err.message };
  }
}
