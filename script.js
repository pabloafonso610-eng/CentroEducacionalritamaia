document.addEventListener('DOMContentLoaded', () => {
  iniciarSlider();
  configurarPrecos();
  configurarFormulario();
  carregarResumoSucesso();
  definirDataMinimaAgendamento();
});

function iniciarSlider() {
  const slides = document.querySelectorAll('.slide');
  if (!slides.length) return;

  let index = 0;
  setInterval(() => {
    slides[index].classList.remove('active');
    index = (index + 1) % slides.length;
    slides[index].classList.add('active');
  }, 4000);
}

function configurarPrecos() {
  const serieSelect = document.getElementById('serie');
  const precoBox = document.getElementById('precoBox');
  if (!serieSelect || !precoBox) return;

  const tabela = {
    'Maternal': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 640,00' },
    '1º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 800,00' },
    '2º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 800,00' },
    '3º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 800,00' },
    '4º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 800,00' },
    '5º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 800,00' },
    '6º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 1.150,00' },
    '7º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 1.150,00' },
    '8º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 1.150,00' },
    '9º Ano': { matricula: 'R$ 350,00', mensalidade: 'R$ 390,00', material: 'R$ 1.150,00' },
    '1º Ano do Ensino Médio': { matricula: 'R$ 350,00', mensalidade: 'R$ 410,00', material: '10x de R$ 149,00' },
    '2º Ano do Ensino Médio': { matricula: 'R$ 350,00', mensalidade: 'R$ 410,00', material: '10x de R$ 149,00' },
    '3º Ano do Ensino Médio': { matricula: 'R$ 350,00', mensalidade: 'R$ 410,00', material: '10x de R$ 149,00' }
  };

  serieSelect.addEventListener('change', () => {
    const dados = tabela[serieSelect.value];
    if (!dados) {
      precoBox.style.display = 'none';
      return;
    }

    document.getElementById('matriculaPreco').textContent = `Taxa de matrícula: ${dados.matricula}`;
    document.getElementById('mensalidadePreco').textContent = `Mensalidade: ${dados.mensalidade}`;
    document.getElementById('materialPreco').textContent = `Material didático: ${dados.material}`;
    precoBox.style.display = 'block';
  });
}

function configurarFormulario() {
  const form = document.getElementById('formMatricula');
  if (!form) return;

  form.addEventListener('submit', (e) => {
    e.preventDefault();

    if (!form.checkValidity()) {
      form.reportValidity();
      return;
    }

    const registro = {
      id: gerarIdSolicitacao(),
      protocolo: gerarProtocolo(),
      nome: obterValor('nome'),
      nascimento: obterValor('nascimento'),
      serie: obterValor('serie'),
      responsavel: obterValor('responsavel'),
      telefone: obterValor('telefone'),
      observacao: obterValor('observacao'),
      agendamentoData: obterValor('agendamentoData'),
      agendamentoHora: obterValor('agendamentoHora'),
      dataEnvio: formatarDataHoraAtual(),
      statusMatricula: 'Pendente',
      portalLiberado: false,
      loginPortal: '',
      senhaPortal: '',
      validadoEm: ''
    };

    try {
      const matriculas = lerListaMatriculas();
      matriculas.push(registro);

      localStorage.setItem('matriculasCERM', JSON.stringify(matriculas));
      localStorage.setItem('protocoloAtual', registro.protocolo);
      localStorage.setItem('matriculaAtual', JSON.stringify(registro));
      sessionStorage.setItem('matriculaAtual', JSON.stringify(registro));

      window.location.href = './sucesso.html';
    } catch (erro) {
      console.error('Erro ao salvar matrícula:', erro);
      alert('Não foi possível enviar a matrícula agora. Verifique se o navegador permite armazenamento local.');
    }
  });
}

function obterValor(id) {
  const campo = document.getElementById(id);
  return campo ? campo.value.trim() : '';
}

function lerListaMatriculas() {
  try {
    const bruto = localStorage.getItem('matriculasCERM');
    if (!bruto) return [];
    const lista = JSON.parse(bruto);
    return Array.isArray(lista) ? lista : [];
  } catch (erro) {
    console.warn('Lista de matrículas corrompida. Reiniciando armazenamento.', erro);
    return [];
  }
}

function definirDataMinimaAgendamento() {
  const campoData = document.getElementById('agendamentoData');
  if (!campoData) return;
  const hoje = new Date();
  campoData.min = hoje.toISOString().split('T')[0];
}

function gerarProtocolo() {
  const numero = Math.floor(100000 + Math.random() * 900000);
  return `CERM-${numero}`;
}

function gerarIdSolicitacao() {
  return `AG-${Date.now().toString().slice(-8)}`;
}

function formatarDataHoraAtual() {
  const agora = new Date();
  return agora.toLocaleString('pt-BR');
}

function formatarData(data) {
  if (!data || !data.includes('-')) return '-';
  const [ano, mes, dia] = data.split('-');
  return `${dia}/${mes}/${ano}`;
}

function carregarResumoSucesso() {
  const protocoloCampo = document.getElementById('protocolo');
  const agendamentoTexto = document.getElementById('agendamentoTexto');
  if (!protocoloCampo) return;

  let matriculaAtual = null;

  try {
    matriculaAtual = JSON.parse(localStorage.getItem('matriculaAtual') || sessionStorage.getItem('matriculaAtual') || 'null');
  } catch (erro) {
    console.warn('Não foi possível ler a matrícula atual.', erro);
  }

  const protocolo = (matriculaAtual && matriculaAtual.protocolo) || localStorage.getItem('protocoloAtual');
  protocoloCampo.textContent = protocolo || 'Protocolo não encontrado';

  if (agendamentoTexto) {
    agendamentoTexto.textContent = matriculaAtual
      ? `ID ${matriculaAtual.id} • ${formatarData(matriculaAtual.agendamentoData)} às ${matriculaAtual.agendamentoHora}`
      : 'Não foi possível carregar o agendamento. Volte e envie a matrícula novamente.';
  }
}

function voltarSite() {
  window.location.href = './index.html';
}

function baixarComprovante() {
  let matriculaAtual = null;

  try {
    matriculaAtual = JSON.parse(localStorage.getItem('matriculaAtual') || sessionStorage.getItem('matriculaAtual') || 'null');
  } catch (erro) {
    console.warn('Erro ao ler comprovante.', erro);
  }

  if (!matriculaAtual) {
    alert('Dados da matrícula não encontrados. Envie a matrícula novamente.');
    return;
  }

  const texto = [
    'CENTRO EDUCACIONAL RITA MAIA',
    '',
    'COMPROVANTE DE MATRÍCULA',
    '',
    `ID da solicitação: ${matriculaAtual.id}`,
    `Protocolo: ${matriculaAtual.protocolo}`,
    `Aluno: ${matriculaAtual.nome}`,
    `Nascimento: ${formatarData(matriculaAtual.nascimento)}`,
    `Série: ${matriculaAtual.serie}`,
    '',
    `Responsável: ${matriculaAtual.responsavel}`,
    `Telefone: ${matriculaAtual.telefone}`,
    `Observação: ${matriculaAtual.observacao || 'Nenhuma'}`,
    '',
    `Agendamento: ${formatarData(matriculaAtual.agendamentoData)} às ${matriculaAtual.agendamentoHora}`,
    `Enviado em: ${matriculaAtual.dataEnvio}`,
    'Status inicial: Pendente de validação da secretaria',
    '',
    'Guarde este comprovante para acompanhar sua matrícula.'
  ].join('\n');

  const blob = new Blob([texto], { type: 'text/plain;charset=utf-8' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `comprovante_${matriculaAtual.protocolo}.txt`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
}
