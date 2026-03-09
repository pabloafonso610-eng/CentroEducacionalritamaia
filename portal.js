document.addEventListener('DOMContentLoaded', () => {
  configurarMostrarSenhaAluno();
  configurarLoginAluno();
  protegerPortalAluno();
  carregarPortalAluno();
});

function configurarMostrarSenhaAluno() {
  const botao = document.getElementById('toggleSenhaAluno');
  const campo = document.getElementById('alunoSenha');
  if (!botao || !campo) return;

  botao.addEventListener('click', () => {
    const visivel = campo.type === 'text';
    campo.type = visivel ? 'password' : 'text';
    botao.setAttribute('aria-pressed', String(!visivel));
    botao.textContent = visivel ? '👁' : '🙈';
  });
}

function configurarLoginAluno() {
  const form = document.getElementById('formLoginAluno');
  if (!form) return;

  form.addEventListener('submit', (e) => {
    e.preventDefault();

    const acesso = (document.getElementById('alunoProtocolo').value || '').trim().toUpperCase();
    const senha = (document.getElementById('alunoSenha').value || '').trim();
    const matriculas = lerDados('matriculasCERM');

    const aluno = matriculas.find((item) => {
      const protocoloItem = (item.protocolo || '').toUpperCase();
      const idItem = (item.id || '').toUpperCase();
      const loginPortal = (item.loginPortal || '').toUpperCase();
      return (acesso === protocoloItem || acesso === idItem || acesso === loginPortal);
    });

    if (!aluno) {
      alert('Matrícula não encontrada. Confira o ID ou protocolo informado.');
      return;
    }

    if (!aluno.portalLiberado || aluno.statusMatricula !== 'Validada') {
      alert('A matrícula ainda não foi validada pela secretaria. Aguarde a liberação do portal.');
      return;
    }

    if ((aluno.senhaPortal || '') !== senha) {
      alert('Senha do portal incorreta.');
      return;
    }

    localStorage.setItem('alunoLogadoCERM', aluno.id);
    window.location.href = 'aluno.html';
  });
}

function protegerPortalAluno() {
  const paginaAtual = (window.location.pathname.split('/').pop() || '').toLowerCase();
  const estaNoPortal = paginaAtual === 'aluno.html';
  if (!estaNoPortal) return;

  const alunoId = localStorage.getItem('alunoLogadoCERM');
  if (!alunoId) {
    window.location.href = 'portal-aluno.html';
  }
}

function carregarPortalAluno() {
  const titulo = document.getElementById('tituloAluno');
  if (!titulo) return;

  const alunoId = localStorage.getItem('alunoLogadoCERM');
  const matriculas = lerDados('matriculasCERM');
  const aluno = matriculas.find((item) => item.id === alunoId);

  if (!aluno) {
    alert('Dados do aluno não encontrados. Faça login novamente.');
    localStorage.removeItem('alunoLogadoCERM');
    window.location.href = 'portal-aluno.html';
    return;
  }

  if (!aluno.portalLiberado || aluno.statusMatricula !== 'Validada') {
    alert('O acesso deste aluno ainda não está liberado pela secretaria.');
    localStorage.removeItem('alunoLogadoCERM');
    window.location.href = 'portal-aluno.html';
    return;
  }

  const calendarios = lerDados('calendarioEscolarCERM')
    .filter((item) => !item.serie || item.serie === 'Todas' || item.serie === aluno.serie)
    .sort((a, b) => (a.data || '').localeCompare(b.data || ''));
  const notas = lerDados('notasCERM').filter((item) => item.alunoId === aluno.id);
  const historico = lerDados('historicoCERM').filter((item) => item.alunoId === aluno.id);
  const horarios = lerDados('horariosCERM').filter((item) => item.alunoId === aluno.id);
  const avaliacoes = lerDados('avaliacoesCERM').filter((item) => item.alunoId === aluno.id)
    .sort((a, b) => `${a.data} ${a.hora}`.localeCompare(`${b.data} ${b.hora}`));
  const mensalidades = lerDados('mensalidadesCERM').filter((item) => item.alunoId === aluno.id)
    .sort((a, b) => (a.vencimento || '').localeCompare(b.vencimento || ''));

  document.getElementById('tituloAluno').textContent = `${aluno.nome} • ${aluno.serie}`;
  document.getElementById('subtituloAluno').textContent = `Matrícula ${aluno.protocolo} • Login ${aluno.loginPortal || aluno.id}`;
  document.getElementById('alunoResumoLateral').textContent = `${aluno.nome} • ${aluno.serie}`;

  preencherCalendario(calendarios);
  preencherNotas(notas);
  preencherHistorico(historico);
  preencherHorarios(horarios);
  preencherAvaliacoes(avaliacoes);
  preencherMensalidades(mensalidades);
  preencherResumo(notas, avaliacoes, mensalidades);

  const btnLogout = document.getElementById('btnLogoutAluno');
  btnLogout?.addEventListener('click', () => {
    localStorage.removeItem('alunoLogadoCERM');
    window.location.href = 'portal-aluno.html';
  });
}

function preencherCalendario(calendarios) {
  const lista = document.getElementById('listaCalendarioAluno');
  if (!lista) return;

  if (!calendarios.length) {
    lista.innerHTML = '<div class="info-card empty-card">Nenhum evento do calendário escolar lançado ainda.</div>';
    return;
  }

  lista.innerHTML = calendarios.map((item) => `
    <article class="info-card">
      <span class="badge">${escaparHtml(item.serie || 'Todas as séries')}</span>
      <h3>${escaparHtml(item.titulo)}</h3>
      <p><strong>Data:</strong> ${formatarData(item.data)}</p>
      <p>${escaparHtml(item.descricao || 'Sem descrição informada.')}</p>
    </article>
  `).join('');
}

function preencherNotas(notas) {
  const lista = document.getElementById('listaNotasAluno');
  if (!lista) return;

  if (!notas.length) {
    lista.innerHTML = '<tr><td colspan="7" class="empty-state">Nenhuma nota cadastrada.</td></tr>';
    document.getElementById('mediaGeralAluno').textContent = '--';
    return;
  }

  lista.innerHTML = notas.map((item) => {
    const media = calcularMedia(item.n1, item.n2, item.n3, item.n4);
    return `
      <tr>
        <td>${escaparHtml(item.disciplina)}</td>
        <td>${item.n1 || '-'}</td>
        <td>${item.n2 || '-'}</td>
        <td>${item.n3 || '-'}</td>
        <td>${item.n4 || '-'}</td>
        <td>${media.toFixed(1)}</td>
        <td>${media >= 7 ? 'Aprovado parcial' : 'Atenção'}</td>
      </tr>
    `;
  }).join('');

  const mediaGeral = notas.reduce((acc, item) => acc + calcularMedia(item.n1, item.n2, item.n3, item.n4), 0) / notas.length;
  document.getElementById('mediaGeralAluno').textContent = mediaGeral.toFixed(1);
}

function preencherHistorico(historico) {
  const lista = document.getElementById('listaHistoricoAluno');
  if (!lista) return;

  if (!historico.length) {
    lista.innerHTML = '<tr><td colspan="4" class="empty-state">Nenhum histórico escolar lançado.</td></tr>';
    return;
  }

  lista.innerHTML = historico.map((item) => `
    <tr>
      <td>${escaparHtml(item.anoLetivo)}</td>
      <td>${escaparHtml(item.serie)}</td>
      <td>${escaparHtml(item.situacao)}</td>
      <td>${escaparHtml(item.observacao || '-')}</td>
    </tr>
  `).join('');
}

function preencherHorarios(horarios) {
  const lista = document.getElementById('listaHorariosAluno');
  if (!lista) return;

  if (!horarios.length) {
    lista.innerHTML = '<tr><td colspan="4" class="empty-state">Nenhuma grade de horário disponível.</td></tr>';
    return;
  }

  lista.innerHTML = horarios.map((item) => `
    <tr>
      <td>${escaparHtml(item.dia)}</td>
      <td>${escaparHtml(item.disciplina)}</td>
      <td>${escaparHtml(item.horario)}</td>
      <td>${escaparHtml(item.professor || '-')}</td>
    </tr>
  `).join('');
}

function preencherAvaliacoes(avaliacoes) {
  const lista = document.getElementById('listaAvaliacoesAluno');
  if (!lista) return;

  if (!avaliacoes.length) {
    lista.innerHTML = '<tr><td colspan="5" class="empty-state">Nenhuma avaliação agendada.</td></tr>';
    document.getElementById('proximaAvaliacaoAluno').textContent = '---';
    return;
  }

  lista.innerHTML = avaliacoes.map((item) => `
    <tr>
      <td>${escaparHtml(item.disciplina)}</td>
      <td>${formatarData(item.data)}</td>
      <td>${escaparHtml(item.hora || '-')}</td>
      <td>${escaparHtml(item.tipo || '-')}</td>
      <td>${escaparHtml(item.conteudo || '-')}</td>
    </tr>
  `).join('');

  const proxima = avaliacoes[0];
  document.getElementById('proximaAvaliacaoAluno').textContent = `${formatarData(proxima.data)} • ${proxima.disciplina}`;
}

function preencherMensalidades(mensalidades) {
  const lista = document.getElementById('listaMensalidadesAluno');
  if (!lista) return;

  if (!mensalidades.length) {
    lista.innerHTML = '<div class="info-card empty-card">Nenhuma mensalidade cadastrada.</div>';
    document.getElementById('proximoVencimentoAluno').textContent = '---';
    return;
  }

  lista.innerHTML = mensalidades.map((item) => `
    <article class="info-card ${item.status === 'Paga' ? 'card-paid' : 'card-pending'}">
      <span class="badge">${escaparHtml(item.status || 'Pendente')}</span>
      <h3>${escaparHtml(item.referencia)}</h3>
      <p><strong>Vencimento:</strong> ${formatarData(item.vencimento)}</p>
      <p><strong>Valor:</strong> ${escaparHtml(item.valor)}</p>
      <p><strong>Mensagem:</strong> ${escaparHtml(item.mensagem || 'Sem mensagem da secretaria.')}</p>
    </article>
  `).join('');

  const proxima = mensalidades.find((item) => item.status !== 'Paga') || mensalidades[0];
  document.getElementById('proximoVencimentoAluno').textContent = `${formatarData(proxima.vencimento)} • ${proxima.valor}`;
}

function preencherResumo(notas, avaliacoes, mensalidades) {
  if (!notas.length) document.getElementById('mediaGeralAluno').textContent = '--';
  if (!avaliacoes.length) document.getElementById('proximaAvaliacaoAluno').textContent = '---';
  if (!mensalidades.length) document.getElementById('proximoVencimentoAluno').textContent = '---';
}

function lerDados(chave) {
  try {
    const bruto = localStorage.getItem(chave);
    if (!bruto) return [];
    const dados = JSON.parse(bruto);
    return Array.isArray(dados) ? dados : [];
  } catch (erro) {
    console.warn('Erro ao ler dados do portal:', chave, erro);
    return [];
  }
}

function calcularMedia(...notas) {
  const valores = notas.map((nota) => Number(String(nota).replace(',', '.'))).filter((nota) => !Number.isNaN(nota));
  if (!valores.length) return 0;
  return valores.reduce((a, b) => a + b, 0) / valores.length;
}

function formatarData(data) {
  if (!data || !data.includes('-')) return data || '-';
  const [ano, mes, dia] = data.split('-');
  return `${dia}/${mes}/${ano}`;
}

function escaparHtml(texto) {
  return String(texto ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
