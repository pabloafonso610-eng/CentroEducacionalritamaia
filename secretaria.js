document.addEventListener('DOMContentLoaded', () => {
  iniciarLogin();
  configurarMostrarSenha();
  protegerPainel();
  inicializarPortalSecretaria();
});

const LOGIN_SECRETARIA = 'CERMSIP';
const SENHA_SECRETARIA = '241417';

function iniciarLogin() {
  const form = document.getElementById('formLoginSecretaria');
  if (!form) return;

  form.addEventListener('submit', (e) => {
    e.preventDefault();

    const usuario = document.getElementById('usuario').value.trim().toUpperCase();
    const senha = document.getElementById('senha').value.trim();

    if (usuario === LOGIN_SECRETARIA && senha === SENHA_SECRETARIA) {
      localStorage.setItem('secretariaLogada', 'true');
      window.location.href = 'secretaria.html';
      return;
    }

    alert('Login ou senha incorretos.');
  });
}

function configurarMostrarSenha() {
  const botao = document.getElementById('toggleSenha');
  const campo = document.getElementById('senha');
  if (!botao || !campo) return;

  botao.addEventListener('click', () => {
    const visivel = campo.type === 'text';
    campo.type = visivel ? 'password' : 'text';
    botao.setAttribute('aria-pressed', String(!visivel));
    botao.textContent = visivel ? '👁' : '🙈';
  });
}

function protegerPainel() {
  const estaNoPainel = window.location.pathname.includes('secretaria.html');
  if (!estaNoPainel) return;

  const logada = localStorage.getItem('secretariaLogada');
  if (logada !== 'true') {
    window.location.href = 'login.html';
  }
}

function inicializarPortalSecretaria() {
  const tabela = document.getElementById('listaMatriculas');
  if (!tabela) return;

  normalizarMatriculas();
  popularSelectsAlunos();
  configurarAcoesPainel();
  configurarFormularios();
  configurarImportacaoExcel();
  renderizarTudo();
}

function normalizarMatriculas() {
  const matriculas = lerDados('matriculasCERM');
  const ajustadas = matriculas.map((item) => ({
    ...item,
    statusMatricula: item.statusMatricula || (item.portalLiberado ? 'Validada' : 'Pendente'),
    portalLiberado: Boolean(item.portalLiberado),
    loginPortal: item.loginPortal || '',
    senhaPortal: item.senhaPortal || '',
    validadoEm: item.validadoEm || ''
  }));
  salvarDados('matriculasCERM', ajustadas);
}

function configurarAcoesPainel() {
  document.getElementById('filtroMatriculas')?.addEventListener('input', renderizarMatriculas);

  document.getElementById('btnLogout')?.addEventListener('click', () => {
    localStorage.removeItem('secretariaLogada');
    window.location.href = 'login.html';
  });

  document.getElementById('btnLimparMatriculas')?.addEventListener('click', () => {
    const confirmar = confirm('Tem certeza que deseja apagar todas as matrículas salvas neste navegador?');
    if (!confirmar) return;

    localStorage.removeItem('matriculasCERM');
    localStorage.removeItem('protocoloAtual');
    localStorage.removeItem('matriculaAtual');
    popularSelectsAlunos();
    renderizarTudo();
  });

  document.getElementById('listaMatriculas')?.addEventListener('click', (e) => {
    const botao = e.target.closest('button[data-acao]');
    if (!botao) return;

    const alunoId = botao.dataset.id;
    const acao = botao.dataset.acao;

    if (acao === 'validar') validarMatricula(alunoId);
    if (acao === 'revogar') revogarMatricula(alunoId);
    if (acao === 'nova-senha') gerarNovaSenha(alunoId);
  });
}

function configurarFormularios() {
  configurarFormularioLista({
    formId: 'formCalendario',
    storageKey: 'calendarioEscolarCERM',
    mapear: () => ({
      id: gerarId('CAL'),
      titulo: valor('calTitulo'),
      data: valor('calData'),
      serie: valor('calSerie') || 'Todas',
      descricao: valor('calDescricao'),
      criadoEm: new Date().toLocaleString('pt-BR')
    }),
    aoSalvar: () => {
      limparFormulario('formCalendario');
      document.getElementById('calSerie').value = 'Todas';
    }
  });

  configurarFormularioLista({
    formId: 'formNotas',
    storageKey: 'notasCERM',
    mapear: () => ({
      id: gerarId('NOTA'),
      alunoId: valor('notaAluno'),
      alunoNome: nomeAlunoPorId(valor('notaAluno')),
      disciplina: valor('notaDisciplina'),
      n1: valor('notaN1'),
      n2: valor('notaN2'),
      n3: valor('notaN3'),
      n4: valor('notaN4'),
      criadoEm: new Date().toLocaleString('pt-BR')
    }),
    dedupe: (lista, novo) => lista.filter((item) => !(item.alunoId === novo.alunoId && item.disciplina.toLowerCase() === novo.disciplina.toLowerCase())),
    aoSalvar: () => limparFormulario('formNotas')
  });

  configurarFormularioLista({
    formId: 'formHistorico',
    storageKey: 'historicoCERM',
    mapear: () => ({
      id: gerarId('HIS'),
      alunoId: valor('historicoAluno'),
      alunoNome: nomeAlunoPorId(valor('historicoAluno')),
      anoLetivo: valor('historicoAno'),
      serie: valor('historicoSerie'),
      situacao: valor('historicoSituacao'),
      observacao: valor('historicoObs')
    }),
    aoSalvar: () => limparFormulario('formHistorico')
  });

  configurarFormularioLista({
    formId: 'formHorarios',
    storageKey: 'horariosCERM',
    mapear: () => ({
      id: gerarId('HOR'),
      alunoId: valor('horarioAluno'),
      alunoNome: nomeAlunoPorId(valor('horarioAluno')),
      dia: valor('horarioDia'),
      disciplina: valor('horarioDisciplina'),
      horario: valor('horarioFaixa'),
      professor: valor('horarioProfessor')
    }),
    aoSalvar: () => limparFormulario('formHorarios')
  });

  configurarFormularioLista({
    formId: 'formAvaliacoes',
    storageKey: 'avaliacoesCERM',
    mapear: () => ({
      id: gerarId('AVA'),
      alunoId: valor('avaliacaoAluno'),
      alunoNome: nomeAlunoPorId(valor('avaliacaoAluno')),
      disciplina: valor('avaliacaoDisciplina'),
      data: valor('avaliacaoData'),
      hora: valor('avaliacaoHora'),
      tipo: valor('avaliacaoTipo'),
      conteudo: valor('avaliacaoConteudo')
    }),
    aoSalvar: () => limparFormulario('formAvaliacoes')
  });

  configurarFormularioLista({
    formId: 'formMensalidades',
    storageKey: 'mensalidadesCERM',
    mapear: () => ({
      id: gerarId('MEN'),
      alunoId: valor('mensalidadeAluno'),
      alunoNome: nomeAlunoPorId(valor('mensalidadeAluno')),
      referencia: valor('mensalidadeReferencia'),
      vencimento: valor('mensalidadeVencimento'),
      valor: valor('mensalidadeValor'),
      status: valor('mensalidadeStatus'),
      mensagem: valor('mensalidadeMensagem')
    }),
    dedupe: (lista, novo) => lista.filter((item) => !(item.alunoId === novo.alunoId && item.referencia.toLowerCase() === novo.referencia.toLowerCase())),
    aoSalvar: () => {
      limparFormulario('formMensalidades');
      document.getElementById('mensalidadeStatus').value = 'Pendente';
    }
  });
}

function configurarFormularioLista({ formId, storageKey, mapear, dedupe, aoSalvar }) {
  const form = document.getElementById(formId);
  if (!form) return;

  form.addEventListener('submit', (e) => {
    e.preventDefault();

    if (!form.checkValidity()) {
      form.reportValidity();
      return;
    }

    const novo = mapear();
    let lista = lerDados(storageKey);
    if (typeof dedupe === 'function') {
      lista = dedupe(lista, novo);
    }
    lista.push(novo);
    salvarDados(storageKey, lista);
    aoSalvar?.();
    renderizarTudo();
  });
}

function popularSelectsAlunos() {
  const matriculas = lerDados('matriculasCERM');
  const selects = ['notaAluno', 'historicoAluno', 'horarioAluno', 'avaliacaoAluno', 'mensalidadeAluno'];

  selects.forEach((id) => {
    const select = document.getElementById(id);
    if (!select) return;

    select.innerHTML = '<option value="">Selecione o aluno</option>';

    if (!matriculas.length) {
      select.innerHTML += '<option value="" disabled>Nenhuma matrícula cadastrada</option>';
      return;
    }

    select.innerHTML += matriculas.map((aluno) => (
      `<option value="${aluno.id}">${escaparHtml(aluno.nome)} • ${escaparHtml(aluno.serie)} • ${escaparHtml(aluno.protocolo)} • ${escaparHtml(aluno.statusMatricula || 'Pendente')}</option>`
    )).join('');
  });
}

function renderizarTudo() {
  renderizarMatriculas();
  renderizarCalendario();
  renderizarNotas();
  renderizarHistorico();
  renderizarHorarios();
  renderizarAvaliacoes();
  renderizarMensalidades();
  atualizarResumo();
}

function renderizarMatriculas() {
  const tabela = document.getElementById('listaMatriculas');
  if (!tabela) return;

  const matriculas = lerDados('matriculasCERM');
  const termo = (document.getElementById('filtroMatriculas')?.value || '').toLowerCase();

  const filtradas = matriculas.filter((item) => {
    const texto = `${item.id} ${item.protocolo} ${item.nome} ${item.serie} ${item.responsavel} ${item.telefone} ${item.observacao} ${item.agendamentoData} ${item.agendamentoHora} ${item.statusMatricula} ${item.loginPortal}`.toLowerCase();
    return texto.includes(termo);
  });

  if (!filtradas.length) {
    tabela.innerHTML = '<tr><td colspan="9" class="empty-state">Nenhuma solicitação encontrada.</td></tr>';
    return;
  }

  tabela.innerHTML = filtradas.slice().reverse().map((item) => {
    const validada = item.statusMatricula === 'Validada';
    const loginPortal = item.loginPortal || item.id;
    const senhaPortal = validada ? item.senhaPortal : 'Aguardando';
    const statusClasse = validada ? 'badge-success' : 'badge-warning';

    return `
      <tr>
        <td>${escaparHtml(item.id)}</td>
        <td>${escaparHtml(item.protocolo)}</td>
        <td>
          <strong>${escaparHtml(item.nome)}</strong><br>
          <small>${escaparHtml(item.responsavel || '-')}</small>
        </td>
        <td>${escaparHtml(item.serie)}</td>
        <td>${formatarData(item.agendamentoData)} ${escaparHtml(item.agendamentoHora || '')}</td>
        <td>
          <span class="table-badge ${statusClasse}">${escaparHtml(item.statusMatricula || 'Pendente')}</span>
          <div class="table-mini">${item.validadoEm ? `Liberado em ${escaparHtml(item.validadoEm)}` : 'Aguardando análise'}</div>
        </td>
        <td>${validada ? escaparHtml(loginPortal) : '<span class="table-mini">Será liberado após validar</span>'}</td>
        <td>${validada ? escaparHtml(senhaPortal) : '<span class="table-mini">Será gerada automaticamente</span>'}</td>
        <td>
          <div class="action-stack">
            <button class="btn btn-primary btn-sm" data-acao="validar" data-id="${escaparHtml(item.id)}">${validada ? 'Revalidar' : 'Validar'}</button>
            <button class="btn btn-dark btn-sm" data-acao="nova-senha" data-id="${escaparHtml(item.id)}" ${validada ? '' : 'disabled'}>Nova senha</button>
            <button class="btn btn-outline btn-sm" data-acao="revogar" data-id="${escaparHtml(item.id)}" ${validada ? '' : 'disabled'}>Revogar</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

function validarMatricula(alunoId) {
  const matriculas = lerDados('matriculasCERM');
  const atualizadas = matriculas.map((item) => {
    if (item.id !== alunoId) return item;

    const senhaPortal = gerarSenhaPortal();
    return {
      ...item,
      statusMatricula: 'Validada',
      portalLiberado: true,
      loginPortal: item.id,
      senhaPortal,
      validadoEm: new Date().toLocaleString('pt-BR')
    };
  });

  salvarDados('matriculasCERM', atualizadas);
  popularSelectsAlunos();
  renderizarTudo();

  const aluno = atualizadas.find((item) => item.id === alunoId);
  if (aluno) {
    alert(`Matrícula validada com sucesso!\n\nAluno: ${aluno.nome}\nLogin do portal: ${aluno.loginPortal}\nSenha do portal: ${aluno.senhaPortal}`);
  }
}

function gerarNovaSenha(alunoId) {
  const matriculas = lerDados('matriculasCERM');
  const atualizadas = matriculas.map((item) => {
    if (item.id !== alunoId || item.statusMatricula !== 'Validada') return item;

    return {
      ...item,
      senhaPortal: gerarSenhaPortal(),
      validadoEm: item.validadoEm || new Date().toLocaleString('pt-BR')
    };
  });

  salvarDados('matriculasCERM', atualizadas);
  renderizarTudo();

  const aluno = atualizadas.find((item) => item.id === alunoId);
  if (aluno) {
    alert(`Nova senha gerada para ${aluno.nome}: ${aluno.senhaPortal}`);
  }
}

function revogarMatricula(alunoId) {
  const confirmar = confirm('Deseja revogar a validação desta matrícula e bloquear o portal do aluno?');
  if (!confirmar) return;

  const matriculas = lerDados('matriculasCERM');
  const atualizadas = matriculas.map((item) => {
    if (item.id !== alunoId) return item;

    return {
      ...item,
      statusMatricula: 'Pendente',
      portalLiberado: false,
      loginPortal: '',
      senhaPortal: '',
      validadoEm: ''
    };
  });

  salvarDados('matriculasCERM', atualizadas);
  popularSelectsAlunos();
  renderizarTudo();
}

function renderizarCalendario() {
  const lista = document.getElementById('listaCalendarioSecretaria');
  if (!lista) return;

  const eventos = lerDados('calendarioEscolarCERM').sort((a, b) => (a.data || '').localeCompare(b.data || ''));
  if (!eventos.length) {
    lista.innerHTML = '<div class="info-card empty-card">Nenhum evento cadastrado no calendário escolar.</div>';
    return;
  }

  lista.innerHTML = eventos.map((item) => `
    <article class="info-card">
      <span class="badge">${escaparHtml(item.serie || 'Todas')}</span>
      <h3>${escaparHtml(item.titulo)}</h3>
      <p><strong>Data:</strong> ${formatarData(item.data)}</p>
      <p>${escaparHtml(item.descricao || 'Sem descrição.')}</p>
    </article>
  `).join('');
}

function renderizarNotas() {
  const lista = document.getElementById('listaNotasSecretaria');
  if (!lista) return;

  const notas = lerDados('notasCERM');
  if (!notas.length) {
    lista.innerHTML = '<tr><td colspan="4" class="empty-state">Nenhuma nota lançada.</td></tr>';
    return;
  }

  lista.innerHTML = notas.slice().reverse().map((item) => `
    <tr>
      <td>${escaparHtml(item.alunoNome)}</td>
      <td>${escaparHtml(item.disciplina)}</td>
      <td>${calcularMedia(item.n1, item.n2, item.n3, item.n4).toFixed(1)}</td>
      <td>${escaparHtml(item.criadoEm || '-')}</td>
    </tr>
  `).join('');
}

function renderizarHistorico() {
  const lista = document.getElementById('listaHistoricoSecretaria');
  if (!lista) return;

  const historico = lerDados('historicoCERM');
  if (!historico.length) {
    lista.innerHTML = '<tr><td colspan="4" class="empty-state">Nenhum histórico lançado.</td></tr>';
    return;
  }

  lista.innerHTML = historico.slice().reverse().map((item) => `
    <tr>
      <td>${escaparHtml(item.alunoNome)}</td>
      <td>${escaparHtml(item.anoLetivo)}</td>
      <td>${escaparHtml(item.serie)}</td>
      <td>${escaparHtml(item.situacao)}</td>
    </tr>
  `).join('');
}

function renderizarHorarios() {
  const lista = document.getElementById('listaHorariosSecretaria');
  if (!lista) return;

  const horarios = lerDados('horariosCERM');
  if (!horarios.length) {
    lista.innerHTML = '<tr><td colspan="4" class="empty-state">Nenhum horário lançado.</td></tr>';
    return;
  }

  lista.innerHTML = horarios.slice().reverse().map((item) => `
    <tr>
      <td>${escaparHtml(item.alunoNome)}</td>
      <td>${escaparHtml(item.dia)}</td>
      <td>${escaparHtml(item.disciplina)}</td>
      <td>${escaparHtml(item.horario)}</td>
    </tr>
  `).join('');
}

function renderizarAvaliacoes() {
  const lista = document.getElementById('listaAvaliacoesSecretaria');
  if (!lista) return;

  const avaliacoes = lerDados('avaliacoesCERM').sort((a, b) => `${a.data} ${a.hora}`.localeCompare(`${b.data} ${b.hora}`));
  if (!avaliacoes.length) {
    lista.innerHTML = '<tr><td colspan="5" class="empty-state">Nenhuma avaliação lançada.</td></tr>';
    return;
  }

  lista.innerHTML = avaliacoes.map((item) => `
    <tr>
      <td>${escaparHtml(item.alunoNome)}</td>
      <td>${escaparHtml(item.disciplina)}</td>
      <td>${formatarData(item.data)}</td>
      <td>${escaparHtml(item.hora)}</td>
      <td>${escaparHtml(item.tipo)}</td>
    </tr>
  `).join('');
}

function renderizarMensalidades() {
  const lista = document.getElementById('listaMensalidadesSecretaria');
  if (!lista) return;

  const mensalidades = lerDados('mensalidadesCERM').sort((a, b) => (a.vencimento || '').localeCompare(b.vencimento || ''));
  if (!mensalidades.length) {
    lista.innerHTML = '<div class="info-card empty-card">Nenhuma mensalidade cadastrada.</div>';
    return;
  }

  lista.innerHTML = mensalidades.map((item) => `
    <article class="info-card ${item.status === 'Paga' ? 'card-paid' : 'card-pending'}">
      <span class="badge">${escaparHtml(item.status)}</span>
      <h3>${escaparHtml(item.alunoNome)} • ${escaparHtml(item.referencia)}</h3>
      <p><strong>Vencimento:</strong> ${formatarData(item.vencimento)}</p>
      <p><strong>Valor:</strong> ${escaparHtml(item.valor)}</p>
      <p>${escaparHtml(item.mensagem || 'Sem mensagem personalizada.')}</p>
    </article>
  `).join('');
}

function atualizarResumo() {
  const matriculas = lerDados('matriculasCERM');
  const mensalidades = lerDados('mensalidadesCERM').sort((a, b) => (a.vencimento || '').localeCompare(b.vencimento || ''));
  const validadas = matriculas.filter((item) => item.statusMatricula === 'Validada');

  document.getElementById('totalMatriculas').textContent = matriculas.length;
  document.getElementById('totalValidadas').textContent = validadas.length;
  document.getElementById('ultimoProtocolo').textContent = matriculas.length ? matriculas[matriculas.length - 1].protocolo : '---';
  document.getElementById('proximoAgendamento').textContent = matriculas.length
    ? `${formatarData(matriculas[matriculas.length - 1].agendamentoData)} ${matriculas[matriculas.length - 1].agendamentoHora || ''}`.trim()
    : '---';

  const proximaMensalidade = mensalidades.find((item) => item.status !== 'Paga') || mensalidades[0];
  document.getElementById('resumoMensalidade').textContent = proximaMensalidade
    ? `${formatarData(proximaMensalidade.vencimento)} • ${proximaMensalidade.valor}`
    : '---';
}

function valor(id) {
  return (document.getElementById(id)?.value || '').trim();
}

function limparFormulario(id) {
  document.getElementById(id)?.reset();
}

function nomeAlunoPorId(alunoId) {
  const aluno = lerDados('matriculasCERM').find((item) => item.id === alunoId);
  return aluno ? aluno.nome : 'Aluno não encontrado';
}

function lerDados(chave) {
  try {
    const bruto = localStorage.getItem(chave);
    if (!bruto) return [];
    const dados = JSON.parse(bruto);
    return Array.isArray(dados) ? dados : [];
  } catch (erro) {
    console.warn('Erro ao ler dados:', chave, erro);
    return [];
  }
}

function salvarDados(chave, dados) {
  localStorage.setItem(chave, JSON.stringify(dados));
}

function gerarId(prefixo) {
  return `${prefixo}-${Date.now().toString().slice(-8)}`;
}

function gerarSenhaPortal() {
  return `AL${Math.floor(1000 + Math.random() * 9000)}`;
}

function formatarData(data) {
  if (!data) return '-';
  if (!data.includes('-')) return data;
  const [ano, mes, dia] = data.split('-');
  return `${dia}/${mes}/${ano}`;
}

function calcularMedia(...notas) {
  const valores = notas.map((nota) => Number(String(nota).replace(',', '.'))).filter((nota) => !Number.isNaN(nota));
  if (!valores.length) return 0;
  return valores.reduce((a, b) => a + b, 0) / valores.length;
}

function escaparHtml(texto) {
  return String(texto ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}


function configurarImportacaoExcel() {
  const botao = document.getElementById('btnImportarExcel');
  const arquivoInput = document.getElementById('arquivoExcel');
  if (!botao || !arquivoInput) return;

  botao.addEventListener('click', async () => {
    const arquivo = arquivoInput.files?.[0];
    if (!arquivo) {
      atualizarFeedbackImportacao('Selecione uma planilha do Excel antes de importar.', true);
      return;
    }

    if (typeof XLSX === 'undefined') {
      atualizarFeedbackImportacao('A biblioteca de leitura do Excel não carregou. Verifique a internet e tente novamente.', true);
      return;
    }

    try {
      const buffer = await arquivo.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
      const resumo = importarWorkbookExcel(workbook);
      popularSelectsAlunos();
      renderizarTudo();
      arquivoInput.value = '';
      atualizarFeedbackImportacao(resumo.mensagem, false);
    } catch (erro) {
      console.error('Erro ao importar planilha:', erro);
      atualizarFeedbackImportacao('Não foi possível ler a planilha. Confira se o arquivo está em .xlsx, .xls ou .csv.', true);
    }
  });
}

function importarWorkbookExcel(workbook) {
  const resumo = {
    matriculas: 0,
    calendario: 0,
    notas: 0,
    historico: 0,
    horarios: 0,
    avaliacoes: 0,
    mensalidades: 0,
    ignoradas: 0,
    totalAbas: workbook.SheetNames.length,
    mensagem: ''
  };

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const linhas = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const aba = identificarAba(sheetName);
    if (!aba || !linhas.length) {
      if (linhas.length) resumo.ignoradas += linhas.length;
      return;
    }

    if (aba === 'matriculas') resumo.matriculas += importarMatriculasExcel(linhas);
    if (aba === 'calendario') resumo.calendario += importarCalendarioExcel(linhas);
    if (aba === 'notas') resumo.notas += importarNotasExcel(linhas);
    if (aba === 'historico') resumo.historico += importarHistoricoExcel(linhas);
    if (aba === 'horarios') resumo.horarios += importarHorariosExcel(linhas);
    if (aba === 'avaliacoes') resumo.avaliacoes += importarAvaliacoesExcel(linhas);
    if (aba === 'mensalidades') resumo.mensalidades += importarMensalidadesExcel(linhas);
  });

  resumo.mensagem = `Importação concluída. Matrículas: ${resumo.matriculas}, calendário: ${resumo.calendario}, notas: ${resumo.notas}, histórico: ${resumo.historico}, horários: ${resumo.horarios}, avaliações: ${resumo.avaliacoes}, mensalidades: ${resumo.mensalidades}. Linhas ignoradas: ${resumo.ignoradas}.`;
  return resumo;
}

function identificarAba(sheetName) {
  const nome = normalizarChave(sheetName);
  if (nome.includes('matricula')) return 'matriculas';
  if (nome.includes('calendario')) return 'calendario';
  if (nome.includes('nota')) return 'notas';
  if (nome.includes('historico')) return 'historico';
  if (nome.includes('horario') || nome.includes('grade')) return 'horarios';
  if (nome.includes('avaliacao')) return 'avaliacoes';
  if (nome.includes('mensalidade') || nome.includes('financeiro')) return 'mensalidades';
  return '';
}

function importarMatriculasExcel(linhas) {
  const lista = lerDados('matriculasCERM');
  let total = 0;

  linhas.forEach((linha) => {
    const row = normalizarLinha(linha);
    const nome = obterValorLinha(row, ['nome', 'aluno', 'nomealuno']);
    if (!nome) return;

    const protocoloBase = obterValorLinha(row, ['protocolo', 'matricula', 'codigo']) || gerarId('MAT');
    const idBase = obterValorLinha(row, ['id', 'idaluno', 'idportal']) || protocoloBase;
    const status = obterValorLinha(row, ['statusmatricula', 'status', 'situacao']) || 'Pendente';
    const loginPortal = obterValorLinha(row, ['loginportal', 'login', 'usuariportal', 'usuarioportal']) || (status.toLowerCase() === 'validada' ? idBase : '');
    const senhaPortal = obterValorLinha(row, ['senhaportal', 'senha']) || '';
    const portalLiberado = textoBooleano(obterValorLinha(row, ['portalliberado', 'liberado'])) || status.toLowerCase() === 'validada';
    const indice = lista.findIndex((item) => item.id === idBase || item.protocolo === protocoloBase || normalizarChave(item.nome) === normalizarChave(nome));

    const base = {
      id: idBase,
      protocolo: protocoloBase,
      nome,
      serie: obterValorLinha(row, ['serie', 'turma']) || 'Não informada',
      responsavel: obterValorLinha(row, ['responsavel', 'nomeresponsavel']) || '',
      telefone: obterValorLinha(row, ['telefone', 'celular', 'fone']) || '',
      nascimento: normalizarData(obterValorLinha(row, ['nascimento', 'datanascimento'])),
      observacao: obterValorLinha(row, ['observacao', 'obs']) || '',
      agendamentoData: normalizarData(obterValorLinha(row, ['agendamentodata', 'dataagendamento'])),
      agendamentoHora: normalizarHora(obterValorLinha(row, ['agendamentohora', 'horaagendamento'])),
      statusMatricula: portalLiberado ? 'Validada' : status,
      portalLiberado,
      loginPortal,
      senhaPortal,
      validadoEm: obterValorLinha(row, ['validadoem', 'datavalidacao']) || (portalLiberado ? new Date().toLocaleString('pt-BR') : '')
    };

    if (indice >= 0) {
      lista[indice] = { ...lista[indice], ...base };
    } else {
      lista.push(base);
    }
    total += 1;
  });

  salvarDados('matriculasCERM', lista);
  return total;
}

function importarCalendarioExcel(linhas) {
  let lista = lerDados('calendarioEscolarCERM');
  let total = 0;
  linhas.forEach((linha) => {
    const row = normalizarLinha(linha);
    const titulo = obterValorLinha(row, ['titulo', 'evento', 'nomeevento']);
    if (!titulo) return;
    const item = {
      id: obterValorLinha(row, ['id']) || gerarId('CAL'),
      titulo,
      data: normalizarData(obterValorLinha(row, ['data', 'dataevento'])),
      serie: obterValorLinha(row, ['serie', 'publico']) || 'Todas',
      descricao: obterValorLinha(row, ['descricao', 'obs', 'observacao']) || '',
      criadoEm: obterValorLinha(row, ['criadoem']) || new Date().toLocaleString('pt-BR')
    };
    lista = lista.filter((atual) => !(normalizarChave(atual.titulo) === normalizarChave(item.titulo) && atual.data === item.data && atual.serie === item.serie));
    lista.push(item);
    total += 1;
  });
  salvarDados('calendarioEscolarCERM', lista);
  return total;
}

function importarNotasExcel(linhas) {
  let lista = lerDados('notasCERM');
  let total = 0;
  linhas.forEach((linha) => {
    const row = normalizarLinha(linha);
    const aluno = localizarAlunoImportacao(row);
    const disciplina = obterValorLinha(row, ['disciplina', 'materia']);
    if (!aluno || !disciplina) return;
    const item = {
      id: obterValorLinha(row, ['id']) || gerarId('NOTA'),
      alunoId: aluno.id,
      alunoNome: aluno.nome,
      disciplina,
      n1: normalizarNumeroTexto(obterValorLinha(row, ['n1', 'nota1', '1nota'])),
      n2: normalizarNumeroTexto(obterValorLinha(row, ['n2', 'nota2', '2nota'])),
      n3: normalizarNumeroTexto(obterValorLinha(row, ['n3', 'nota3', '3nota'])),
      n4: normalizarNumeroTexto(obterValorLinha(row, ['n4', 'nota4', '4nota'])),
      criadoEm: obterValorLinha(row, ['criadoem', 'lancamento']) || new Date().toLocaleString('pt-BR')
    };
    lista = lista.filter((atual) => !(atual.alunoId === item.alunoId && normalizarChave(atual.disciplina) === normalizarChave(item.disciplina)));
    lista.push(item);
    total += 1;
  });
  salvarDados('notasCERM', lista);
  return total;
}

function importarHistoricoExcel(linhas) {
  let lista = lerDados('historicoCERM');
  let total = 0;
  linhas.forEach((linha) => {
    const row = normalizarLinha(linha);
    const aluno = localizarAlunoImportacao(row);
    const anoLetivo = obterValorLinha(row, ['anoletivo', 'ano']);
    if (!aluno || !anoLetivo) return;
    const item = {
      id: obterValorLinha(row, ['id']) || gerarId('HIS'),
      alunoId: aluno.id,
      alunoNome: aluno.nome,
      anoLetivo,
      serie: obterValorLinha(row, ['serie', 'seriecursada']) || aluno.serie,
      situacao: obterValorLinha(row, ['situacao', 'status']) || 'Em andamento',
      observacao: obterValorLinha(row, ['observacao', 'obs']) || ''
    };
    lista = lista.filter((atual) => !(atual.alunoId === item.alunoId && String(atual.anoLetivo) === String(item.anoLetivo) && normalizarChave(atual.serie) === normalizarChave(item.serie)));
    lista.push(item);
    total += 1;
  });
  salvarDados('historicoCERM', lista);
  return total;
}

function importarHorariosExcel(linhas) {
  let lista = lerDados('horariosCERM');
  let total = 0;
  linhas.forEach((linha) => {
    const row = normalizarLinha(linha);
    const aluno = localizarAlunoImportacao(row);
    const disciplina = obterValorLinha(row, ['disciplina', 'materia']);
    if (!aluno || !disciplina) return;
    const item = {
      id: obterValorLinha(row, ['id']) || gerarId('HOR'),
      alunoId: aluno.id,
      alunoNome: aluno.nome,
      dia: obterValorLinha(row, ['dia', 'diasemana']) || '',
      disciplina,
      horario: obterValorLinha(row, ['horario', 'faixa', 'hora']) || '',
      professor: obterValorLinha(row, ['professor', 'docente']) || ''
    };
    lista = lista.filter((atual) => !(atual.alunoId === item.alunoId && normalizarChave(atual.disciplina) === normalizarChave(item.disciplina) && normalizarChave(atual.dia) === normalizarChave(item.dia) && normalizarChave(atual.horario) === normalizarChave(item.horario)));
    lista.push(item);
    total += 1;
  });
  salvarDados('horariosCERM', lista);
  return total;
}

function importarAvaliacoesExcel(linhas) {
  let lista = lerDados('avaliacoesCERM');
  let total = 0;
  linhas.forEach((linha) => {
    const row = normalizarLinha(linha);
    const aluno = localizarAlunoImportacao(row);
    const disciplina = obterValorLinha(row, ['disciplina', 'materia']);
    if (!aluno || !disciplina) return;
    const item = {
      id: obterValorLinha(row, ['id']) || gerarId('AVA'),
      alunoId: aluno.id,
      alunoNome: aluno.nome,
      disciplina,
      data: normalizarData(obterValorLinha(row, ['data', 'dataavaliacao'])),
      hora: normalizarHora(obterValorLinha(row, ['hora', 'horario'])),
      tipo: obterValorLinha(row, ['tipo', 'tipoavaliacao']) || '',
      conteudo: obterValorLinha(row, ['conteudo', 'observacao', 'obs']) || ''
    };
    lista = lista.filter((atual) => !(atual.alunoId === item.alunoId && normalizarChave(atual.disciplina) === normalizarChave(item.disciplina) && atual.data === item.data && atual.hora === item.hora));
    lista.push(item);
    total += 1;
  });
  salvarDados('avaliacoesCERM', lista);
  return total;
}

function importarMensalidadesExcel(linhas) {
  let lista = lerDados('mensalidadesCERM');
  let total = 0;
  linhas.forEach((linha) => {
    const row = normalizarLinha(linha);
    const aluno = localizarAlunoImportacao(row);
    const referencia = obterValorLinha(row, ['referencia', 'mesreferencia', 'competencia']);
    if (!aluno || !referencia) return;
    const item = {
      id: obterValorLinha(row, ['id']) || gerarId('MEN'),
      alunoId: aluno.id,
      alunoNome: aluno.nome,
      referencia,
      vencimento: normalizarData(obterValorLinha(row, ['vencimento', 'datavencimento'])),
      valor: obterValorLinha(row, ['valor', 'valorcobrado']) || '',
      status: obterValorLinha(row, ['status', 'situacao']) || 'Pendente',
      mensagem: obterValorLinha(row, ['mensagem', 'observacao', 'obs']) || ''
    };
    lista = lista.filter((atual) => !(atual.alunoId === item.alunoId && normalizarChave(atual.referencia) === normalizarChave(item.referencia)));
    lista.push(item);
    total += 1;
  });
  salvarDados('mensalidadesCERM', lista);
  return total;
}

function localizarAlunoImportacao(row) {
  const lista = lerDados('matriculasCERM');
  const id = obterValorLinha(row, ['alunoid', 'id', 'idaluno', 'idportal']);
  const protocolo = obterValorLinha(row, ['protocolo', 'matricula', 'codigo']);
  const nome = obterValorLinha(row, ['aluno', 'nome', 'nomealuno']);

  return lista.find((item) => {
    const bateId = id && (item.id === id || item.loginPortal === id);
    const bateProtocolo = protocolo && item.protocolo === protocolo;
    const bateNome = nome && normalizarChave(item.nome) === normalizarChave(nome);
    return bateId || bateProtocolo || bateNome;
  }) || null;
}

function atualizarFeedbackImportacao(texto, erro = false) {
  const caixa = document.getElementById('resultadoImportacao');
  if (!caixa) return;
  caixa.textContent = texto;
  caixa.style.borderColor = erro ? '#ef4444' : '#22c55e';
  caixa.style.color = erro ? '#b91c1c' : '#166534';
}

function normalizarLinha(linha) {
  return Object.entries(linha).reduce((acc, [chave, valor]) => {
    acc[normalizarChave(chave)] = valor;
    return acc;
  }, {});
}

function obterValorLinha(row, aliases) {
  for (const alias of aliases) {
    const valor = row[normalizarChave(alias)];
    if (valor !== undefined && valor !== null && String(valor).trim() !== '') {
      return String(valor).trim();
    }
  }
  return '';
}

function normalizarChave(texto) {
  return String(texto || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function normalizarData(valor) {
  const texto = String(valor || '').trim();
  if (!texto) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
    const [dia, mes, ano] = texto.split('/');
    return `${ano}-${mes}-${dia}`;
  }
  if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(texto)) {
    const [dia, mes, ano] = texto.split('-').map((p) => p.padStart(2, '0'));
    return `${ano}-${mes}-${dia}`;
  }
  const numero = Number(texto);
  if (!Number.isNaN(numero) && numero > 20000 && numero < 60000) {
    const base = new Date(Date.UTC(1899, 11, 30));
    base.setUTCDate(base.getUTCDate() + numero);
    return base.toISOString().slice(0, 10);
  }
  return texto;
}

function normalizarHora(valor) {
  const texto = String(valor || '').trim();
  if (!texto) return '';
  if (/^\d{2}:\d{2}$/.test(texto)) return texto;
  const numero = Number(texto);
  if (!Number.isNaN(numero) && numero >= 0 && numero < 1) {
    const totalMin = Math.round(numero * 24 * 60);
    const h = String(Math.floor(totalMin / 60)).padStart(2, '0');
    const m = String(totalMin % 60).padStart(2, '0');
    return `${h}:${m}`;
  }
  return texto;
}

function normalizarNumeroTexto(valor) {
  const texto = String(valor || '').trim();
  if (!texto) return '';
  return texto.replace('.', ',');
}

function textoBooleano(valor) {
  const texto = normalizarChave(valor);
  return ['sim', 'true', '1', 'validada', 'liberado'].includes(texto);
}
