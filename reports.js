import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

export function formatDateBr(isoOrDate) {
  if (!isoOrDate) return "-";
  try {
    const d = typeof isoOrDate === "string" ? new Date(isoOrDate) : isoOrDate;
    if (isNaN(d.getTime())) return String(isoOrDate);
    return d.toLocaleDateString("pt-BR", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    });
  } catch {
    return String(isoOrDate);
  }
}

export function formatDateTimeBr(d = new Date()) {
  return d.toLocaleDateString("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

export function getSituacaoEspecial(item) {
  const status = item.status || "Ativo";
  const interesse = item.interesse || "Médio";
  if (status === "Pronto para batismo" || status === "Batismo Realizado") {
    return "Decisão / Batismo";
  }
  if (status === "Pronto para apelo") {
    return "Pronto para Apelo";
  }
  if (status === "Pausado" || status === "Desinteressado" || interesse === "Baixo") {
    return "Atenção / Em Risco";
  }
  return "Em Andamento Regular";
}

/**
 * Universal Data Resolver: Guarantees 100% of student fields (study name, progress, instructors, contacts, notes)
 * are cleanly populated regardless of legacy or new schema differences.
 */
export function resolveInteressadoFullData(item, state = {}) {
  if (!item) return {};

  const usuarios = state.usuarios || [];
  const series = state.series || [];
  const locais = state.locais || [];

  // 1. Nome
  const nome = (item.nome || item.name || "Sem Nome").trim();

  // 2. Contato
  const telefone = (item.telefone || item.celular || item.whatsapp || item.contato || item.fone || "").trim();

  // 3. Endereço
  const endereco = (item.endereco || item.logradouro || item.rua || item.bairro || "").trim();

  // 4. Igreja / Local
  let igrejaNome = (item.igrejaNome || item.localNome || item.local || item.igreja || "").trim();
  if (!igrejaNome && item.igrejaId) {
    const loc = locais.find((l) => l.id === item.igrejaId);
    if (loc) igrejaNome = loc.nome;
  }
  if (!igrejaNome && state.user?.igrejaNome) {
    igrejaNome = state.user.igrejaNome;
  }
  if (!igrejaNome) igrejaNome = "Sem igreja";

  const igrejaTipo = item.igrejaTipo || (igrejaNome.toLowerCase().includes("grupo") ? "grupo" : "igreja");
  const distrito = item.distrito || "Esplanada";

  // 5. Série Bíblica ("Estudo que está")
  let serieNome = (item.serieNome || item.guiaEstudo || item.curso || item.estudoNome || item.estudo || item.serie || "").trim();
  let matchedSerie = null;
  if (item.serieId) {
    matchedSerie = series.find((s) => s.id === item.serieId);
  }
  if (!matchedSerie && serieNome) {
    matchedSerie = series.find((s) => s.nome?.toLowerCase() === serieNome.toLowerCase());
  }
  if (matchedSerie) {
    serieNome = matchedSerie.nome;
  }
  if (!serieNome) {
    serieNome = series[0]?.nome || "Série Bíblica";
  }

  // 6. Lições e Progresso
  let totalEstudos = Number(item.totalEstudos || item.totalLicoes || item.total || 0);
  if (totalEstudos <= 0 && matchedSerie) {
    totalEstudos = Number(matchedSerie.totalEstudos || matchedSerie.total || 0);
  }
  if (totalEstudos <= 0) {
    totalEstudos = 18; // default fallback
  }

  const rawAtual = item.estudoAtual ?? item.licaoAtual ?? item.licao ?? item.estudoNumero ?? 0;
  const estudoAtual = Math.max(0, Number(rawAtual || 0));
  const capped = Math.min(estudoAtual, totalEstudos);
  const porcentagem = Math.round((capped / totalEstudos) * 100);
  const faltantes = Math.max(0, totalEstudos - capped);
  const progressStr = `${porcentagem}% (${capped}/${totalEstudos})`;

  // 7. Instrutores ("Instrutor da pessoa")
  const instNamesSet = new Set();

  if (Array.isArray(item.instrutorNomes) && item.instrutorNomes.length) {
    item.instrutorNomes.forEach((n) => {
      if (n && typeof n === "string" && n.trim()) instNamesSet.add(n.trim());
    });
  }

  if (Array.isArray(item.instrutores) && item.instrutores.length) {
    item.instrutores.forEach((i) => {
      if (typeof i === "object" && i?.nome) instNamesSet.add(i.nome.trim());
      else if (typeof i === "string" && i.trim()) instNamesSet.add(i.trim());
    });
  }

  if (Array.isArray(item.instrutorIds) && item.instrutorIds.length) {
    item.instrutorIds.forEach((id) => {
      const u = usuarios.find((user) => (user.id || user.uid) === id);
      if (u?.nome) instNamesSet.add(u.nome.trim());
    });
  }

  if (item.responsavelNome && item.responsavelNome.trim()) {
    instNamesSet.add(item.responsavelNome.trim());
  } else if (item.responsavelId) {
    const u = usuarios.find((user) => (user.id || user.uid) === item.responsavelId);
    if (u?.nome) instNamesSet.add(u.nome.trim());
  }

  if (instNamesSet.size === 0 && item.criadoPorNome && item.criadoPorNome.trim()) {
    instNamesSet.add(item.criadoPorNome.trim());
  } else if (instNamesSet.size === 0 && (item.criadoPorId || item.criadoPorUid)) {
    const u = usuarios.find((user) => (user.id || user.uid) === (item.criadoPorId || item.criadoPorUid));
    if (u?.nome) instNamesSet.add(u.nome.trim());
  }

  const instrutoresList = Array.from(instNamesSet);
  const instrutoresString = instrutoresList.length > 0 ? instrutoresList.join(", ") : "Não atribuído";

  // 8. Status & Interesse
  const status = item.status || "Ativo";
  const interesse = item.interesse || "Médio";
  const observacoes = (item.observacoes || item.notas || item.obs || "").trim();
  const ultimoContato = item.ultimoContato || item.dataUltimoContato || "";
  const createdAt = item.createdAt || item.dataCadastro || item.criadoEm || "";
  const updatedAt = item.updatedAt || item.ultimaAtualizacao || item.atualizadoEm || "";
  const criadoPorNome = item.criadoPorNome || item.criadorNome || "-";
  const criadoPorPerfil = item.criadoPorPerfil || "-";

  return {
    ...item,
    nome,
    telefone: telefone || "Sem telefone",
    telefoneRaw: telefone,
    endereco: endereco || "Sem endereço cadastrado",
    enderecoRaw: endereco,
    igrejaNome,
    igrejaTipo,
    distrito,
    serieNome,
    totalEstudos,
    estudoAtual: capped,
    porcentagem,
    faltantes,
    progressStr,
    instrutorNomes: instrutoresList,
    instrutoresString,
    status,
    interesse,
    observacoes: observacoes || "(Sem observações pastorais registradas)",
    observacoesRaw: observacoes,
    ultimoContato,
    createdAt,
    updatedAt,
    criadoPorNome,
    criadoPorPerfil,
    situacaoPastoral: getSituacaoEspecial({ status, interesse })
  };
}

export function filterInteressadosForReport(interessados, filters = {}, state = {}) {
  let list = Array.isArray(interessados) ? [...interessados] : [];

  // First resolve full data for every student
  list = list.map((i) => resolveInteressadoFullData(i, state));

  if (filters.igreja && filters.igreja !== "Todas") {
    list = list.filter((i) => i.igrejaNome === filters.igreja || i.igrejaId === filters.igreja);
  }

  if (filters.status && filters.status !== "Todos") {
    list = list.filter((i) => i.status === filters.status);
  }

  if (filters.serie && filters.serie !== "Todas") {
    list = list.filter((i) => i.serieId === filters.serie || i.serieNome === filters.serie);
  }

  if (filters.interesse && filters.interesse !== "Todos") {
    list = list.filter((i) => i.interesse === filters.interesse);
  }

  return list;
}

export function buildInstructorsData(interessados, usuarios, state = {}) {
  const map = new Map();

  // Seed from registered users (profiles)
  (usuarios || []).forEach((u) => {
    map.set(u.id || u.uid, {
      id: u.id || u.uid,
      nome: u.nome || "Sem Nome",
      email: u.email || u.emailAuth || "-",
      igrejaNome: u.igrejaNome || "-",
      perfil: u.perfil || "instrutor",
      students: []
    });
  });

  // Attach students
  (interessados || []).forEach((rawStudent) => {
    const student = resolveInteressadoFullData(rawStudent, state);
    let linked = false;

    if (Array.isArray(student.instrutorIds) && student.instrutorIds.length) {
      student.instrutorIds.forEach((instId) => {
        if (!map.has(instId)) {
          map.set(instId, {
            id: instId,
            nome: student.instrutoresString || "Instrutor",
            email: "-",
            igrejaNome: student.igrejaNome || "-",
            perfil: "instrutor",
            students: []
          });
        }
        map.get(instId).students.push(student);
        linked = true;
      });
    }

    if (!linked && student.responsavelId) {
      if (!map.has(student.responsavelId)) {
        map.set(student.responsavelId, {
          id: student.responsavelId,
          nome: student.responsavelNome || "Instrutor",
          email: "-",
          igrejaNome: student.igrejaNome || "-",
          perfil: "instrutor",
          students: []
        });
      }
      map.get(student.responsavelId).students.push(student);
      linked = true;
    }

    if (!linked && (student.criadoPorId || student.criadoPorUid)) {
      const creatorId = student.criadoPorId || student.criadoPorUid;
      if (map.has(creatorId)) {
        map.get(creatorId).students.push(student);
      }
    }
  });

  return Array.from(map.values()).sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
}

/* =========================================================
   EXCEL EXPORT (Multi-sheet & Clean Structure)
========================================================= */
export function exportExcelReport({ reportType, filters, state }) {
  const wb = XLSX.utils.book_new();
  const filteredInteressados = filterInteressadosForReport(state.interessados, filters, state);
  const nowStr = formatDateTimeBr();

  const addSheetWithColWidths = (data, sheetName) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const colWidths = Object.keys(data[0] || {}).map((key) => {
      let maxLen = key.length;
      data.forEach((row) => {
        const val = row[key] ? String(row[key]) : "";
        if (val.length > maxLen) maxLen = val.length;
      });
      return { wch: Math.min(50, Math.max(12, maxLen + 2)) };
    });
    ws["!cols"] = colWidths;
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  };

  // 1. Aba de Resumo Executivo
  if (reportType === "completo" || reportType === "detalhado" || reportType === "resumo") {
    const totalEstudantes = filteredInteressados.length;
    const ativos = filteredInteressados.filter((i) => i.status === "Ativo" || i.status === "Estudando").length;
    const prontoApelo = filteredInteressados.filter((i) => i.status === "Pronto para apelo").length;
    const prontoBatismo = filteredInteressados.filter((i) => i.status === "Pronto para batismo").length;
    const batizados = filteredInteressados.filter((i) => i.status === "Batismo Realizado" || i.status === "Concluído").length;
    const pausados = filteredInteressados.filter((i) => i.status === "Pausado" || i.status === "Desinteressado").length;

    const resumoRows = [
      { Métrica: "Título do Relatório", Valor: reportType === "detalhado" ? "Dossiê Cadastral Completo - Esplanada Viva" : "Relatório de Gestão - Esplanada Viva" },
      { Métrica: "Data de Emissão", Valor: nowStr },
      { Métrica: "Distrito", Valor: "Esplanada" },
      { Métrica: "Filtro Igreja", Valor: filters.igreja || "Todas" },
      { Métrica: "Filtro Status", Valor: filters.status || "Todos" },
      { Métrica: "Filtro Série", Valor: filters.serie || "Todas" },
      { Métrica: "Filtro Nível de Interesse", Valor: filters.interesse || "Todos" },
      { Métrica: "Total de Estudantes / Interessados", Valor: totalEstudantes },
      { Métrica: "Estudos em Andamento (Ativos)", Valor: ativos },
      { Métrica: "Prontos para Apelo", Valor: prontoApelo },
      { Métrica: "Prontos para Batismo", Valor: prontoBatismo },
      { Métrica: "Batismos Realizados / Concluídos", Valor: batizados },
      { Métrica: "Pausados / Desinteressados", Valor: pausados },
      { Métrica: "Total de Séries Cadastradas", Valor: (state.series || []).length },
      { Métrica: "Total de Locais / Igrejas", Valor: (state.locais || []).length },
      { Métrica: "Total de Usuários Cadastrados", Valor: (state.usuarios || []).length }
    ];
    addSheetWithColWidths(resumoRows, "Resumo Executivo");
  }

  // 2. Aba DETALHADA COM TODOS OS CAMPOS (100% dos dados)
  if (reportType === "detalhado" || reportType === "completo" || reportType === "interessados_detalhado") {
    const detailedRows = filteredInteressados.map((item, index) => ({
      "Nº": index + 1,
      "Nome do Interessado": item.nome,
      "Telefone / WhatsApp": item.telefone,
      "Endereço Completo": item.endereco,
      "Igreja / Local": item.igrejaNome,
      "Tipo de Localidade": item.igrejaTipo === "grupo" ? "Grupo" : "Igreja",
      "Distrito": item.distrito,
      "Status do Estudo": item.status,
      "Nível de Interesse": item.interesse,
      "Série Bíblica": item.serieNome,
      "Lição Atual": item.estudoAtual,
      "Total de Lições": item.totalEstudos,
      "Progresso (%)": `${item.porcentagem}%`,
      "Lições Restantes": item.faltantes,
      "Instrutores Responsáveis": item.instrutoresString,
      "Data Último Contato": formatDateBr(item.ultimoContato),
      "Observações Pastorais / Notas": item.observacoes,
      "Cadastrado Por": item.criadoPorNome,
      "Perfil do Cadastrador": item.criadoPorPerfil,
      "Data de Cadastro": formatDateBr(item.createdAt),
      "Última Atualização": formatDateBr(item.updatedAt),
      "Situação Pastoral": item.situacaoPastoral
    }));

    if (detailedRows.length > 0) {
      addSheetWithColWidths(detailedRows, "Fichas Detalhadas (Completo)");
    } else {
      addSheetWithColWidths([{ Aviso: "Nenhum estudante encontrado para os filtros selecionados." }], "Fichas Detalhadas");
    }
  }

  // 3. Aba de Interessados Sintética (Padrão)
  if (reportType === "interessados") {
    const rows = filteredInteressados.map((item, index) => ({
      "Nº": index + 1,
      "Nome do Interessado": item.nome,
      "Telefone": item.telefone,
      "Endereço": item.endereco,
      "Igreja / Local": item.igrejaNome,
      "Série Bíblica": item.serieNome,
      "Progresso": item.progressStr,
      "Status": item.status,
      "Interesse": item.interesse,
      "Instrutor(es)": item.instrutoresString,
      "Último Contato": formatDateBr(item.ultimoContato),
      "Data de Cadastro": formatDateBr(item.createdAt),
      "Observações": item.observacoes
    }));

    if (rows.length > 0) {
      addSheetWithColWidths(rows, "Estudantes e Interessados");
    } else {
      addSheetWithColWidths([{ Aviso: "Nenhum estudante encontrado para os filtros selecionados." }], "Estudantes");
    }
  }

  // 4. Aba de Instrutores e Alunos
  if (reportType === "completo" || reportType === "detalhado" || reportType === "instrutores") {
    const instructorsList = buildInstructorsData(state.interessados, state.usuarios, state);
    const rows = [];

    instructorsList.forEach((inst) => {
      if (filters.igreja && filters.igreja !== "Todas" && inst.igrejaNome !== filters.igreja) {
        return;
      }
      if (inst.students.length === 0) {
        rows.push({
          "Instrutor": inst.nome,
          "Igreja / Local": inst.igrejaNome,
          "Email": inst.email,
          "Perfil": inst.perfil,
          "Total Alunos": 0,
          "Nome do Aluno": "(Sem alunos no momento)",
          "Status do Aluno": "-",
          "Série do Aluno": "-",
          "Progresso": "-",
          "Último Contato": "-",
          "Notas do Aluno": "-"
        });
      } else {
        inst.students.forEach((st) => {
          rows.push({
            "Instrutor": inst.nome,
            "Igreja / Local": inst.igrejaNome,
            "Email": inst.email,
            "Perfil": inst.perfil,
            "Total Alunos": inst.students.length,
            "Nome do Aluno": st.nome,
            "Status do Aluno": st.status,
            "Série do Aluno": st.serieNome,
            "Progresso": st.progressStr,
            "Último Contato": formatDateBr(st.ultimoContato),
            "Notas do Aluno": st.observacoes
          });
        });
      }
    });

    if (rows.length > 0) {
      addSheetWithColWidths(rows, "Instrutores e Alunos");
    } else {
      addSheetWithColWidths([{ Aviso: "Nenhum instrutor encontrado." }], "Instrutores");
    }
  }

  // 5. Aba de Locais e Igrejas
  if (reportType === "completo" || reportType === "detalhado" || reportType === "locais") {
    const allStudentsResolved = (state.interessados || []).map((i) => resolveInteressadoFullData(i, state));
    const rows = (state.locais || []).map((loc) => {
      const students = allStudentsResolved.filter((i) => i.igrejaNome === loc.nome || i.igrejaId === loc.id);
      const batizados = students.filter((i) => i.status === "Batismo Realizado" || i.status === "Concluído").length;
      const ativos = students.filter((i) => i.status === "Ativo" || i.status === "Estudando").length;
      const prontosBatismo = students.filter((i) => i.status === "Pronto para batismo").length;
      const prontosApelo = students.filter((i) => i.status === "Pronto para apelo").length;

      return {
        "Localidade": loc.nome,
        "Tipo": loc.tipo === "grupo" ? "Grupo" : "Igreja",
        "Total de Estudantes": students.length,
        "Estudos em Andamento (Ativos)": ativos,
        "Prontos para Apelo": prontosApelo,
        "Prontos para Batismo": prontosBatismo,
        "Batismos Realizados": batizados
      };
    });

    if (rows.length > 0) {
      addSheetWithColWidths(rows, "Locais e Igrejas");
    }
  }

  // 6. Aba de Séries
  if (reportType === "completo" || reportType === "detalhado" || reportType === "series") {
    const allStudentsResolved = (state.interessados || []).map((i) => resolveInteressadoFullData(i, state));
    const rows = (state.series || []).map((serie) => {
      const count = allStudentsResolved.filter((i) => i.serieId === serie.id || i.serieNome === serie.nome).length;
      return {
        "Nome da Série": serie.nome,
        "Total de Lições": serie.totalEstudos || serie.total || 0,
        "Tipo": serie.padrao ? "Padrão do Sistema" : "Personalizada",
        "Alunos Cursando": count
      };
    });

    if (rows.length > 0) {
      addSheetWithColWidths(rows, "Catálogo de Séries");
    }
  }

  const dateFile = new Date().toISOString().slice(0, 10);
  const prefix = reportType === "detalhado" ? "Relatorio_Detalhado_Esplanada_Viva" : `Relatorio_Esplanada_Viva_${reportType}`;
  const fileName = `${prefix}_${dateFile}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

/* =========================================================
   PDF EXPORT (High-craft Document with Autotable)
========================================================= */
export function exportPdfReport({ reportType, filters, state }) {
  const isLandscape = true;
  const doc = new jsPDF({
    orientation: isLandscape ? "landscape" : "portrait",
    unit: "mm",
    format: "a4"
  });

  const primaryColor = [24, 121, 78]; // #18794e
  const burgundyColor = [155, 34, 38];
  const darkTextColor = [30, 41, 59];

  const filteredInteressados = filterInteressadosForReport(state.interessados, filters, state);
  const nowStr = formatDateTimeBr();

  const drawHeader = (title, subtitle) => {
    doc.setFillColor(...primaryColor);
    doc.rect(0, 0, doc.internal.pageSize.width, 24, "F");

    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.text("ESPLANADA VIVA - MINISTÉRIO DE ESTUDOS BÍBLICOS", 14, 11);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text("Distrito de Esplanada | 'Ele cuida e me ensina a cuidar!'", 14, 18);

    doc.setFontSize(9);
    doc.text(`Emissão: ${nowStr}`, doc.internal.pageSize.width - 14, 18, { align: "right" });

    // Title box
    doc.setTextColor(...darkTextColor);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.text(title, 14, 32);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(100, 116, 139);
    doc.text(subtitle, 14, 37);
  };

  const getReportTitle = () => {
    switch (reportType) {
      case "detalhado":
      case "interessados_detalhado":
        return {
          title: "RELATÓRIO DETALHADO DE INTERESSADOS (DOSSIÊ COMPLETO)",
          subtitle: `Relação cadastral completa com todos os dados, contatos, progresso, instrutores e notas pastorais | ${filteredInteressados.length} registros`
        };
      case "completo":
        return {
          title: "RELATÓRIO CONSOLIDADO COMPLETO",
          subtitle: `Filtros: Igreja (${filters.igreja || "Todas"}) | Status (${filters.status || "Todos"}) | Série (${filters.serie || "Todas"})`
        };
      case "interessados":
        return {
          title: "RELATÓRIO DE ESTUDANTES E INTERESSADOS",
          subtitle: `Total de registros filtrados: ${filteredInteressados.length} | Igreja: ${filters.igreja || "Todas"}`
        };
      case "instrutores":
        return {
          title: "RELATÓRIO DE INSTRUTORES E SEUS ALUNOS",
          subtitle: `Acompanhamento pastoral de instrutores e discípulos`
        };
      case "locais":
        return {
          title: "RELATÓRIO DE IGREJAS E LOCAIS DO DISTRITO",
          subtitle: `Distribuição de interessados e batismos por localidade`
        };
      case "series":
        return {
          title: "RELATÓRIO DO CATÁLOGO DE ESTUDOS BÍBLICOS",
          subtitle: `Estatísticas de adesão às séries e lições`
        };
      case "resumo":
      default:
        return {
          title: "RESUMO EXECUTIVO E INDICADORES",
          subtitle: `Visão geral do ministério e decisões`
        };
    }
  };

  const { title, subtitle } = getReportTitle();
  drawHeader(title, subtitle);

  let currentY = 42;

  // =======================================================
  // CASO: RELATÓRIO DETALHADO (Tabela Completa com 100% dos dados)
  // =======================================================
  if (reportType === "detalhado" || reportType === "interessados_detalhado") {
    const tableBody = filteredInteressados.map((item, idx) => [
      `${idx + 1}`,
      `${item.nome}\nTel: ${item.telefone}\nEnd: ${item.endereco}`,
      `${item.igrejaNome}\nDistrito: ${item.distrito}`,
      `${item.serieNome}\nLição: ${item.estudoAtual}/${item.totalEstudos} (${item.porcentagem}%)\nFaltam: ${item.faltantes}`,
      `Status: ${item.status}\nInteresse: ${item.interesse}\nÚlt. Contato: ${formatDateBr(item.ultimoContato)}`,
      `Instrutor(es):\n${item.instrutoresString}\n\nCadastrado por:\n${item.criadoPorNome} (${formatDateBr(item.createdAt)})`,
      item.observacoes
    ]);

    autoTable(doc, {
      startY: currentY,
      head: [[
        "#",
        "Estudante & Contato",
        "Localidade",
        "Estudo Bíblico & Progresso",
        "Situação & Contato",
        "Instrutores & Cadastro",
        "Observações / Notas Pastorais"
      ]],
      body: tableBody.length > 0 ? tableBody : [["-", "Nenhum estudante encontrado", "-", "-", "-", "-", "-"]],
      theme: "grid",
      headStyles: { fillColor: primaryColor, textColor: 255, fontStyle: "bold", fontSize: 8 },
      styles: { fontSize: 7.5, cellPadding: 2.5, overflow: "linebreak" },
      columnStyles: {
        0: { cellWidth: 8, halign: "center", fontStyle: "bold" },
        1: { cellWidth: 48, fontStyle: "bold" },
        2: { cellWidth: 32 },
        3: { cellWidth: 40 },
        4: { cellWidth: 38 },
        5: { cellWidth: 45 },
        6: { cellWidth: 58 }
      }
    });

    currentY = doc.lastAutoTable.finalY + 10;
  }

  // 1. Resumo Executivo / Métricas
  else if (reportType === "resumo" || reportType === "completo") {
    const total = filteredInteressados.length;
    const ativos = filteredInteressados.filter((i) => i.status === "Ativo" || i.status === "Estudando").length;
    const prontoBatismo = filteredInteressados.filter((i) => i.status === "Pronto para batismo").length;
    const batizados = filteredInteressados.filter((i) => i.status === "Batismo Realizado" || i.status === "Concluído").length;
    const pausados = filteredInteressados.filter((i) => i.status === "Pausado" || i.status === "Desinteressado").length;

    autoTable(doc, {
      startY: currentY,
      head: [["Indicador / Métrica", "Quantidade / Valor", "Status / Detalhes"]],
      body: [
        ["Total de Estudantes Cadastrados", `${total}`, "Base ativa no distrito"],
        ["Estudos em Andamento (Ativos)", `${ativos}`, "Recebendo acompanhamento contínuo"],
        ["Prontos para Batismo / Decisão", `${prontoBatismo}`, "Foco para apelo e colheita"],
        ["Batismos Realizados / Concluídos", `${batizados}`, "Frutos do ministério"],
        ["Estudos Pausados ou Desinteressados", `${pausados}`, "Necessitam de visita ou reativação"],
        ["Total de Igrejas / Locais", `${(state.locais || []).length}`, "Distrito de Esplanada"],
        ["Total de Séries no Catálogo", `${(state.series || []).length}`, "Cursos bíblicos disponíveis"]
      ],
      theme: "striped",
      headStyles: { fillColor: primaryColor, textColor: 255, fontStyle: "bold" },
      styles: { fontSize: 9, cellPadding: 3 },
      columnStyles: {
        0: { fontStyle: "bold", cellWidth: 80 },
        1: { cellWidth: 50, halign: "center", fontStyle: "bold", textColor: burgundyColor }
      }
    });

    currentY = doc.lastAutoTable.finalY + 10;
  }

  // 2. Tabela de Estudantes / Interessados (Padrão ou parte do Completo)
  if (reportType === "interessados" || (reportType === "completo" && currentY)) {
    if (reportType === "completo") {
      doc.addPage();
      drawHeader("RELAÇÃO DE ESTUDANTES E INTERESSADOS", `Filtro aplicado: ${filteredInteressados.length} estudantes listados`);
      currentY = 42;
    }

    const tableBody = filteredInteressados.map((item, idx) => [
      `${idx + 1}`,
      item.nome,
      item.telefone,
      item.igrejaNome,
      item.serieNome,
      item.progressStr,
      item.status,
      item.interesse,
      item.instrutoresString
    ]);

    autoTable(doc, {
      startY: currentY,
      head: [["#", "Nome do Estudante", "Telefone", "Igreja", "Série", "Progresso", "Status", "Interesse", "Instrutor(es)"]],
      body: tableBody.length > 0 ? tableBody : [["-", "Nenhum estudante encontrado", "-", "-", "-", "-", "-", "-", "-"]],
      theme: "grid",
      headStyles: { fillColor: primaryColor, textColor: 255, fontStyle: "bold", fontSize: 8 },
      styles: { fontSize: 7.5, cellPadding: 2, overflow: "linebreak" },
      columnStyles: {
        0: { cellWidth: 8, halign: "center" },
        1: { cellWidth: 40, fontStyle: "bold" },
        2: { cellWidth: 26 },
        3: { cellWidth: 32 },
        4: { cellWidth: 35 },
        5: { cellWidth: 22, halign: "center" },
        6: { cellWidth: 28 },
        7: { cellWidth: 18, halign: "center" },
        8: { cellWidth: 45 }
      }
    });

    currentY = doc.lastAutoTable.finalY + 10;
  }

  // 3. Tabela de Instrutores e Alunos
  if (reportType === "instrutores") {
    const instructorsList = buildInstructorsData(state.interessados, state.usuarios, state);
    const tableBody = [];

    instructorsList.forEach((inst) => {
      if (filters.igreja && filters.igreja !== "Todas" && inst.igrejaNome !== filters.igreja) return;

      if (inst.students.length === 0) {
        tableBody.push([
          inst.nome,
          inst.igrejaNome || "-",
          inst.email || "-",
          "0",
          "(Sem alunos no momento)",
          "-",
          "-"
        ]);
      } else {
        inst.students.forEach((st, sIdx) => {
          tableBody.push([
            sIdx === 0 ? inst.nome : "",
            sIdx === 0 ? inst.igrejaNome || "-" : "",
            sIdx === 0 ? inst.email || "-" : "",
            sIdx === 0 ? `${inst.students.length}` : "",
            st.nome,
            st.status || "Ativo",
            `${st.serieNome} (${st.progressStr})`
          ]);
        });
      }
    });

    autoTable(doc, {
      startY: currentY,
      head: [["Instrutor", "Igreja", "Contato", "Total", "Estudante Vinculado", "Status", "Série & Progresso"]],
      body: tableBody.length > 0 ? tableBody : [["-", "-", "-", "-", "Nenhum instrutor encontrado", "-", "-"]],
      theme: "grid",
      headStyles: { fillColor: primaryColor, textColor: 255, fontStyle: "bold", fontSize: 8 },
      styles: { fontSize: 7.5, cellPadding: 2.2 },
      columnStyles: {
        0: { cellWidth: 40, fontStyle: "bold" },
        1: { cellWidth: 32 },
        2: { cellWidth: 38 },
        3: { cellWidth: 15, halign: "center", fontStyle: "bold" },
        4: { cellWidth: 45, fontStyle: "bold" },
        5: { cellWidth: 30 },
        6: { cellWidth: 50 }
      }
    });
  }

  // 4. Locais
  if (reportType === "locais") {
    const allStudentsResolved = (state.interessados || []).map((i) => resolveInteressadoFullData(i, state));
    const tableBody = (state.locais || []).map((loc) => {
      const students = allStudentsResolved.filter((i) => i.igrejaNome === loc.nome || i.igrejaId === loc.id);
      const batizados = students.filter((i) => i.status === "Batismo Realizado" || i.status === "Concluído").length;
      const ativos = students.filter((i) => i.status === "Ativo" || i.status === "Estudando").length;
      return [
        loc.nome,
        loc.tipo === "grupo" ? "Grupo" : "Igreja",
        `${students.length}`,
        `${ativos}`,
        `${batizados}`
      ];
    });

    autoTable(doc, {
      startY: currentY,
      head: [["Igreja / Grupo Local", "Tipo", "Total de Estudantes", "Estudando (Ativos)", "Batismos / Concluídos"]],
      body: tableBody.length > 0 ? tableBody : [["-", "-", "-", "-", "-"]],
      theme: "striped",
      headStyles: { fillColor: primaryColor, textColor: 255, fontStyle: "bold" },
      styles: { fontSize: 8.5, cellPadding: 3 },
      columnStyles: {
        0: { fontStyle: "bold" },
        2: { halign: "center", fontStyle: "bold" },
        3: { halign: "center" },
        4: { halign: "center", fontStyle: "bold", textColor: burgundyColor }
      }
    });
  }

  // 5. Séries
  if (reportType === "series") {
    const allStudentsResolved = (state.interessados || []).map((i) => resolveInteressadoFullData(i, state));
    const tableBody = (state.series || []).map((serie) => {
      const count = allStudentsResolved.filter((i) => i.serieId === serie.id || i.serieNome === serie.nome).length;
      return [
        serie.nome,
        `${serie.totalEstudos || serie.total || 0} lições`,
        serie.padrao ? "Série Padrão" : "Personalizada",
        `${count} estudantes`
      ];
    });

    autoTable(doc, {
      startY: currentY,
      head: [["Nome da Série de Estudo", "Total de Lições", "Classificação", "Alunos Matriculados"]],
      body: tableBody.length > 0 ? tableBody : [["-", "-", "-", "-"]],
      theme: "striped",
      headStyles: { fillColor: primaryColor, textColor: 255, fontStyle: "bold" },
      styles: { fontSize: 8.5, cellPadding: 3 },
      columnStyles: {
        0: { fontStyle: "bold" },
        1: { halign: "center" },
        2: { halign: "center" },
        3: { halign: "center", fontStyle: "bold", textColor: primaryColor }
      }
    });
  }

  // Footers with page number
  const pageCount = doc.internal.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(140, 150, 160);
    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    doc.text(
      `Esplanada Viva © ${new Date().getFullYear()} - Sistema de Gestão de Estudos Bíblicos | Página ${i} de ${pageCount}`,
      pageWidth / 2,
      pageHeight - 8,
      { align: "center" }
    );
  }

  const dateFile = new Date().toISOString().slice(0, 10);
  const prefix = reportType === "detalhado" ? "Relatorio_Detalhado_Esplanada_Viva" : `Relatorio_Esplanada_Viva_${reportType}`;
  const fileName = `${prefix}_${dateFile}.pdf`;
  doc.save(fileName);
}
