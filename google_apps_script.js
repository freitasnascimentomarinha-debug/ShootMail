// ══════════════════════════════════════════════════════════════════
//  ShootMail — Google Apps Script Backend  v3.1
//  Cole este código em: Planilha → Extensões → Apps Script
// ══════════════════════════════════════════════════════════════════

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ── Nomes das abas ──
const ABA = {
    fornecedores: 'Fornecedores',
    processos: 'Processos',
    destinatarios: 'Destinatarios_Processo',
    disparos: 'Disparos',
    respostas: 'Respostas',
    config: 'Config',
    templates: 'Templates',
    autoDispatch: 'AutoDispatch',
};

// ══════════════════════════════════════════════════════════════════
//  SETUP — cria as abas e o menu na 1ª execução
// ══════════════════════════════════════════════════════════════════
function setupPlanilha() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const abas = {
        [ABA.fornecedores]: [
            'ID', 'Nome', 'Email', 'Email2', 'Email3', 'Tipo_Material', 'Materiais',
            'Telefone', 'CNPJ', 'Cidade_UF', 'Status', 'Observacoes', 'Total_Enviados', 'Criado_Em', 'Codigo_Item'
        ],
        [ABA.processos]: [
            'ID', 'Assunto', 'Corpo_Email', 'Criado_Em', 'Status', 'Total_Destinatarios',
            'Total_Disparos', 'Total_Abriram', 'Total_Responderam', 'Agendado_Para', 'Atualizado_Em', 'Observacoes'
        ],
        [ABA.destinatarios]: [
            'ID', 'Processo_ID', 'Fornecedor_ID', 'Nome_Fornecedor', 'Email_Fornecedor',
            'Tipo_Material', 'Abriu', 'Abriu_Em', 'Respondeu', 'Total_Disparos_Individual'
        ],
        [ABA.disparos]: [
            'ID', 'Processo_ID', 'Fornecedor_ID', 'Nome_Fornecedor', 'Email_Fornecedor',
            'Data_Hora', 'Nota', 'Status_Envio', 'Erro'
        ],
        [ABA.respostas]: [
            'ID', 'Processo_ID', 'Fornecedor_ID', 'Nome_Fornecedor', 'Email_Fornecedor',
            'Assunto_Resposta', 'Preview', 'Respondido_Em', 'Link_Email', 'Anexos', 'Processo_Assunto'
        ],
        [ABA.config]: ['Chave', 'Valor'],
        [ABA.templates]: ['ID', 'Nome', 'Categoria', 'Assunto', 'Corpo'],
        [ABA.autoDispatch]: ['ProcessoID', 'IntervaloMs', 'MaxReenvios', 'ReenviosFeitos', 'ProximoEnvio', 'Ativo'],
    };

    for (const [nome, headers] of Object.entries(abas)) {
        let sheet = ss.getSheetByName(nome);
        if (!sheet) {
            sheet = ss.insertSheet(nome);
            sheet.appendRow(headers);
            sheet.getRange(1, 1, 1, headers.length).setBackground('#1c1f2e').setFontColor('#6b6f90').setFontWeight('bold');
            sheet.setFrozenRows(1);
        } else {
            // Verifica e adiciona colunas faltantes
            const currentHeaders = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
            headers.forEach((h, i) => {
                if (!currentHeaders.includes(h)) {
                    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h)
                        .setBackground('#1c1f2e').setFontColor('#6b6f90').setFontWeight('bold');
                }
            });
        }
    }

    // Configs padrão
    const cfg = ss.getSheetByName(ABA.config);
    const existingKeys = cfg.getLastRow() > 1
        ? cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().flat()
        : [];
    const defaults = [
        ['gmail_remetente', ''],
        ['gmail_nome', 'ShootMail'],
        ['verificacao_ativa', 'false'],
    ];
    defaults.forEach(([k, v]) => {
        if (!existingKeys.includes(k)) cfg.appendRow([k, v]);
    });

    criarMenu();
    SpreadsheetApp.getUi().alert('✅ ShootMail configurado com sucesso!\n\nAs abas foram criadas. Agora publique como Web App.');
}

function criarMenu() {
    SpreadsheetApp.getUi()
        .createMenu('⚡ ShootMail')
        .addItem('🔧 Configurar Planilha', 'setupPlanilha')
        .addItem('📬 Verificar Gmail Agora', 'verificarRespostas')
        .addItem('▶ Ativar Verificação Automática (15min)', 'ativarTrigger')
        .addItem('⏹ Desativar Verificação Automática', 'desativarTrigger')
        .addToUi();
}

function onOpen() { criarMenu(); }

// ══════════════════════════════════════════════════════════════════
//  WEB APP — entry points
//  IMPORTANTE: Usar GET evita erros de CORS preflight no browser.
//  O frontend envia: ?action=xxx&payload={...json...}
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
    const action = (e.parameter && e.parameter.action || '').trim().toLowerCase();

    // ── RASTREAMENTO (PIXEL) ──
    if (action === 'track') return handleTracking(e.parameter);

    // ── DOWNLOAD DIRETO (Evita limites de JSONP) ──
    if (action === 'baixar_anexo' && e.parameter.dl === '1') {
        try {
            const res = baixarAnexo(e.parameter.msgId, e.parameter.filename);
            const html = `
          <html>
            <body style="font-family:sans-serif; text-align:center; padding-top:50px; background:#080a10; color:#fff;">
              <p>Baixando <b>${res.name}</b>...</p>
              <script>
                (function() {
                  const b64 = "${res.content}";
                  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
                  const blob = new Blob([bytes], { type: "${res.contentType}" });
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement('a');
                  a.href = url; a.download = "${res.name}";
                  document.body.appendChild(a); a.click();
                  setTimeout(() => window.close(), 2000);
                })();
              </script>
            </body>
          </html>
        `;
            return HtmlService.createHtmlOutput(html).setTitle('Baixando anexo...');
        } catch (err) {
            return HtmlService.createHtmlOutput('<b>Erro ao baixar:</b> ' + err.toString());
        }
    }

    const output = _processRequest(e.parameter);
    const json = JSON.stringify(output);
    const callback = e.parameter && e.parameter.callback;

    if (callback) {
        // JSONP — funciona sem erro de CORS mesmo em file://
        return ContentService
            .createTextOutput(callback + '(' + json + ')')
            .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
        .createTextOutput(json)
        .setMimeType(ContentService.MimeType.JSON);
}

// Mantido para compatibilidade com chamadas POST legadas
function doPost(e) {
    try {
        const body = JSON.parse(e.postData.contents);
        const output = _processRequest(body);
        return ContentService
            .createTextOutput(JSON.stringify(output))
            .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
        return ContentService
            .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

function _processRequest(params) {
    try {
        const action = (params.action || '').trim().toLowerCase();
        // Suporte a payload JSON (enviado pelo frontend via GET)
        let body = {};
        if (params.payload) {
            try { body = JSON.parse(params.payload); } catch (e) { body = {}; }
        } else {
            body = params; // fallback para parâmetros planos
        }

        let result;
        switch (action) {
            // ── Fornecedores ──
            case 'listar_fornecedores': result = listarFornecedores(); break;
            case 'salvar_fornecedor': result = salvarFornecedor(body.data); break;
            case 'deletar_fornecedor': result = deletarFornecedor(body.id); break;
            case 'toggle_status_fornecedor': result = toggleStatusFornecedor(body.id); break;

            // ── Processos ──
            case 'listar_processos': result = listarProcessos(body.incluir_destinatarios); break;
            case 'salvar_processo': result = salvarProcesso(body.data); break;

            // ── Email ──
            case 'enviar_email': result = enviarEmail(body.data); break;

            // ── Respostas ──
            case 'todas_respostas': result = todasRespostas(); break;
            case 'verificar_respostas': result = verificarRespostas(); break;
            case 'baixar_anexo': result = baixarAnexo(body.msgId, body.filename); break;
            case 'ativar_auto_sync': result = ativarTrigger(); break;
            case 'desativar_auto_sync': result = desativarTrigger(); break;

            // ── Config ──
            case 'ler_todas_configs': result = lerTodasConfigs(); break;
            case 'salvar_config': result = salvarConfig(body.chave, body.valor); break;
            case 'listar_disparos': result = listarDisparos(body.procId); break;

            // ── Rastreamento ──
            case 'track': return handleTracking(params);

            // ── Obs Processo ──
            case 'salvar_obs_processo': result = salvarObsProcesso(body.pid, body.obs); break;

            // ── Templates ──
            case 'salvar_template': result = salvarTemplate(body); break;
            case 'listar_templates': result = listarTemplates(); break;
            case 'deletar_template': result = deletarTemplate(body.id); break;
            case 'atualizar_template': result = atualizarTemplate(body); break;

            // ── Auto-dispatch ──
            case 'configurar_auto_dispatch': result = configurarAutoDispatch(body); break;
            case 'ativar_auto_dispatch_trigger': result = ativarTriggerAutoDispatch(); break;

            default:
                // Ping de status
                return { ok: true, msg: 'ShootMail API ativa', action: action || 'ping' };
        }
        return { ok: true, data: result };
    } catch (err) {
        return { ok: false, error: err.toString() };
    }
}
// ══════════════════════════════════════════════════════════════════
//  FORNECEDORES
// ══════════════════════════════════════════════════════════════════
function listarFornecedores() {
    const sheet = getSheet(ABA.fornecedores);
    return sheetToObjects(sheet);
}

function salvarFornecedor(d) {
    const sheet = getSheet(ABA.fornecedores);
    const now = new Date().toISOString();

    if (d.id) {
        // Editar existente
        const row = findRow(sheet, d.id);
        if (!row) throw new Error('Fornecedor não encontrado: ' + d.id);
        const vals = sheet.getRange(row, 1, 1, 15).getValues()[0];
        sheet.getRange(row, 1, 1, 15).setValues([[
            vals[0],                        // ID (mantém)
            d.nome !== undefined ? d.nome : vals[1],
            d.email !== undefined ? d.email : vals[2],
            d.email2 !== undefined ? d.email2 : vals[3],
            d.email3 !== undefined ? d.email3 : vals[4],
            d.tipo !== undefined ? d.tipo : vals[5],
            d.materiais !== undefined ? d.materiais : vals[6],
            d.telefone !== undefined ? d.telefone : vals[7],
            d.cnpj !== undefined ? d.cnpj : vals[8],
            d.cidade !== undefined ? d.cidade : vals[9],
            d.status !== undefined ? d.status : vals[10],
            d.notas !== undefined ? d.notas : vals[11],
            vals[12],                       // Total_Enviados (mantém)
            vals[13],                       // Criado_Em (mantém)
            d.codigo_item !== undefined ? d.codigo_item : vals[14]
        ]]);
        return { id: d.id };
    } else {
        // Novo
        const id = nextId(sheet);
        sheet.appendRow([
            id, d.nome, d.email,
            d.email2 || '', d.email3 || '',
            d.tipo, d.materiais,
            d.telefone || '', d.cnpj || '', d.cidade || '',
            d.status || 'active',
            d.notas || '', 0, now,
            d.codigo_item || ''
        ]);
        return { id };
    }
}

function deletarFornecedor(id) {
    const sheet = getSheet(ABA.fornecedores);
    const row = findRow(sheet, id);
    if (row) sheet.deleteRow(row);
    return { ok: true };
}

function toggleStatusFornecedor(id) {
    const sheet = getSheet(ABA.fornecedores);
    const row = findRow(sheet, id);
    if (!row) throw new Error('Fornecedor não encontrado');
    const statusCol = 11; // coluna Status (1-indexed)
    const cur = sheet.getRange(row, statusCol).getValue();
    const novo = cur === 'active' ? 'inactive' : 'active';
    sheet.getRange(row, statusCol).setValue(novo);
    return { novo_status: novo };
}

// ══════════════════════════════════════════════════════════════════
//  PROCESSOS
// ══════════════════════════════════════════════════════════════════
function salvarProcesso(d) {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const procSheet = getSheet(ABA.processos);
    const destSheet = getSheet(ABA.destinatarios);
    const now = new Date().toISOString();
    const id = d.id || nextId(procSheet);

    procSheet.appendRow([
        id, d.assunto, d.corpo,
        now,
        d.status || 'active',
        d.total_destinatarios || 0,
        0, 0, 0,
        d.agendado_para || '',
        now
    ]);

    // Grava destinatários em lote, se fornecidos
    if (d.destinatarios && Array.from(d.destinatarios).length > 0) {
        let lastDestId = nextId(destSheet);
        const rows = d.destinatarios.map(rec => [
            lastDestId++,
            id,
            rec.sid,
            rec.name,
            rec.email,
            rec.type || '',
            false, // Abriu
            '',    // Abriu_Em
            false, // Respondeu
            0      // Total_Disparos_Individual (Começa em 0 para o primeiro envio somar 1)
        ]);
        destSheet.getRange(destSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    SpreadsheetApp.flush();
    return { id };
}

function listarProcessos(incluirDest) {
    const sheet = getSheet(ABA.processos);
    const procs = sheetToObjects(sheet);

    // O frontend do ShootMail espera os destinatários embutidos para exibição correta
    // Vamos sempre incluir se houver processos
    if (procs.length > 0) {
        const destSheet = getSheet(ABA.destinatarios);
        const dests = sheetToObjects(destSheet);

        return procs.map(p => ({
            ...p,
            destinatarios: dests.filter(d => String(d.Processo_ID) === String(p.ID))
        }));
    }
    return procs;
}

function salvarObsProcesso(pid, obs) {
    const sheet = getSheet(ABA.processos);
    const row = findRow(sheet, pid);
    if (!row) throw new Error('Processo não encontrado: ' + pid);
    sheet.getRange(row, 12).setValue(obs);
    return { ok: true };
}

// ══════════════════════════════════════════════════════════════════
//  EMAIL
// ══════════════════════════════════════════════════════════════════
function enviarEmail(d) {
    const now = new Date().toISOString();
    const nome = d.nome_remetente || 'ShootMail';
    let statusEnvio = 'enviado';
    let erro = '';

    // ID único para threading inquebrável
    const smId = `[SM-ID:${d.processo_id}-${d.fornecedor_id}]`;
    const assuntoComId = d.assunto + " " + smId;

    // URL de Rastreamento (Pixel)
    const scriptUrl = ScriptApp.getService().getUrl();
    const pixelUrl = `${scriptUrl}?action=track&pid=${d.processo_id}&fid=${d.fornecedor_id}&v=${Date.now()}`;
    const pixelHtml = `<img src="${pixelUrl}" width="1" height="1" style="display:none !important;" />`;


    // Resolve variáveis no corpo
    const corpo = (d.corpo || '')
        .replace(/{nome}/g, d.nome_fornecedor || '')
        .replace(/{empresa}/g, d.nome_fornecedor || '')
        .replace(/{tipo}/g, d.tipo_material || '');

    // Lista de emails para enviar (suporte a email2 / email3)
    const emails = [d.para, d.para2, d.para3].filter(Boolean);

    // Suporte a anexos (array de {name, type, data})
    const attachments = (d.anexos || []).map(file => {
        try {
            const bytes = Utilities.base64Decode(file.data);
            return Utilities.newBlob(bytes, file.type, file.name);
        } catch (e) {
            console.error('Erro ao processar anexo:', file.name, e);
            return null;
        }
    }).filter(Boolean);

    try {
        const htmlBody = corpo.replace(/\n/g, '<br>') + '<br><br>' + pixelHtml;
        emails.forEach(emailDest => {
            GmailApp.sendEmail(emailDest, assuntoComId, corpo, {
                name: nome,
                htmlBody: htmlBody,
                attachments: attachments
            });
        });
    } catch (err) {
        statusEnvio = 'erro';
        erro = err.toString();
    }

    // Registra na aba Disparos
    const disparosSheet = getSheet(ABA.disparos);
    const idD = nextId(disparosSheet);
    disparosSheet.appendRow([
        idD,
        d.processo_id,
        d.fornecedor_id,
        d.nome_fornecedor,
        d.para,
        now,
        d.nota || 'Envio',
        statusEnvio,
        erro
    ]);

    // Cria/atualiza destinatário no processo
    const destSheet = getSheet(ABA.destinatarios);
    const dests = sheetToObjects(destSheet);
    const existe = dests.find(x =>
        String(x.Processo_ID) === String(d.processo_id) &&
        String(x.Fornecedor_ID) === String(d.fornecedor_id)
    );

    if (existe) {
        // Incrementa contador de disparos individuais
        const row = findRowByMultiple(destSheet, [
            { col: 2, val: d.processo_id },
            { col: 3, val: d.fornecedor_id }
        ]);
        if (row) {
            const cur = parseInt(destSheet.getRange(row, 10).getValue()) || 0;
            destSheet.getRange(row, 10).setValue(cur + 1);
        }
    } else {
        const idDest = nextId(destSheet);
        destSheet.appendRow([
            idDest,
            d.processo_id,
            d.fornecedor_id,
            d.nome_fornecedor,
            d.para,
            d.tipo_material || '',
            false, '', false, 1
        ]);
    }
    SpreadsheetApp.flush();

    // Atualiza contador Total_Disparos no processo
    atualizarContadoresProcesso(d.processo_id);

    // Atualiza Total_Enviados no fornecedor
    const fSheet = getSheet(ABA.fornecedores);
    const fRow = findRow(fSheet, d.fornecedor_id);
    if (fRow) {
        const cur = parseInt(fSheet.getRange(fRow, 13).getValue()) || 0;
        fSheet.getRange(fRow, 13).setValue(cur + 1);
    }

    return { id: idD, status: statusEnvio, erro };
}

function listarDisparos(procId) {
    const sheet = getSheet(ABA.disparos);
    const data = sheetToObjects(sheet);
    if (!procId) return data;
    return data.filter(d => String(d.Processo_ID) === String(procId));
}

// ══════════════════════════════════════════════════════════════════
//  RESPOSTAS — verifica caixa de entrada do Gmail
// ══════════════════════════════════════════════════════════════════
function verificarRespostas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const procSheet = ss.getSheetByName(ABA.processos);
    const destSheet = ss.getSheetByName(ABA.destinatarios);
    const respSheet = ss.getSheetByName(ABA.respostas);
    const fornSheet = ss.getSheetByName(ABA.fornecedores);
    if (!procSheet || !destSheet || !respSheet || !fornSheet) return { novas_respostas: 0 };

    const processos = sheetToObjects(procSheet);
    const destinatariosRows = sheetToObjects(destSheet);
    const respostas = sheetToObjects(respSheet);
    const fornecedores = sheetToObjects(fornSheet);

    const registrados = new Set(respostas.map(r => r.Link_Email).filter(Boolean));
    const mainEmail = getMyEmailSafely();
    const aliases = [mainEmail, ...GmailApp.getAliases()].filter(Boolean).map(a => a.toLowerCase());
    let novas = 0;

    // Busca apenas threads recentes para performance
    const threads = GmailApp.search('newer_than:30d');

    threads.forEach(thread => {
        const msgs = thread.getMessages();
        msgs.forEach(msg => {
            const msgId = msg.getId();
            if (registrados.has(msgId)) return;

            const fromEmail = extractEmail(msg.getFrom()).toLowerCase();

            // 1. NUNCA conta emails enviados por VOCÊ mesmo (ou seus aliases) como resposta
            if (aliases.includes(fromEmail)) return;

            const subjectOrig = msg.getSubject();

            // Tenta extrair ID único do assunto [SM-ID:PROCESSO-FORNECEDOR]
            const smMatch = subjectOrig.match(/\[SM-ID:(\d+)-(\d+)\]/);
            let procAlvo = null;
            let destAlvo = null;
            let forn = null;

            if (smMatch) {
                const pid = smMatch[1];
                const fid = smMatch[2];
                procAlvo = processos.find(p => String(p.ID) === String(pid));
                forn = fornecedores.find(f => String(f.ID) === String(fid));

                if (procAlvo && forn) {
                    // 2. Valida se o REMETENTE é realmente o fornecedor (ou um de seus emails)
                    const isForn = [forn.Email, forn.Email2, forn.Email3].filter(Boolean)
                        .some(e => e.toLowerCase() === fromEmail.toLowerCase());

                    if (isForn) {
                        destAlvo = destinatariosRows.find(d => String(d.Processo_ID) === String(pid) && String(d.Fornecedor_ID) === String(fid));
                    } else {
                        // Se o ID bate mas o email não, pode ser um encaminhamento ou erro. 
                        // Vamos resetar para tentar a lógica de fallback por domínio/email se necessário.
                        procAlvo = null;
                        forn = null;
                    }
                }
            }

            if (!procAlvo) {
                // Fallback para lógica legada (assunto + email)
                const subjectClean = subjectOrig.replace(/^(Re|Res|Fwd|Enc|Respv):\s*/i, '').toLowerCase().trim();

                // 1. Identifica fornecedor
                forn = fornecedores.find(f =>
                    f.Email === fromEmail || f.Email2 === fromEmail || f.Email3 === fromEmail
                );
                if (!forn) return;

                // 2. Busca todos os vínculos desse fornecedor
                const vinculos = destinatariosRows.filter(d => String(d.Fornecedor_ID) === String(forn.ID));
                if (!vinculos.length) return;

                // 3. Tenta encontrar o processo que melhor casa com o assunto
                for (const d of vinculos) {
                    const p = processos.find(x => String(x.ID) === String(d.Processo_ID));
                    if (!p) continue;

                    const pSub = (p.Assunto || '').toLowerCase().trim();
                    if (subjectClean.includes(pSub.slice(0, 20)) || pSub.includes(subjectClean.slice(0, 20))) {
                        procAlvo = p;
                        destAlvo = d;
                        break;
                    }
                }

                // Fallback: se não casou assunto, pega o vínculo mais recente
                if (!procAlvo) {
                    destAlvo = vinculos[vinculos.length - 1];
                    procAlvo = processos.find(x => String(x.ID) === String(destAlvo.Processo_ID));
                }
            }

            if (!procAlvo || !forn) return;

            const anexos = msg.getAttachments().map(a => a.getName()).join('; ');
            const date = msg.getDate().toISOString();

            // Registra resposta
            const idR = nextId(respSheet);
            respSheet.appendRow([
                idR, procAlvo.ID, forn.ID, forn.Nome, fromEmail,
                subjectOrig, msg.getPlainBody().slice(0, 500), date, msgId, anexos, procAlvo.Assunto
            ]);
            registrados.add(msgId);
            novas++;

            // Marca destinatário como 'Respondeu' E 'Abriu' (garantia de rastreio)
            const rowDest = findRowByMultiple(destSheet, [
                { col: 2, val: procAlvo.ID },
                { col: 3, val: forn.ID }
            ]);
            if (rowDest) {
                destSheet.getRange(rowDest, 7).setValue(true); // Abriu
                if (!destSheet.getRange(rowDest, 8).getValue()) {
                    destSheet.getRange(rowDest, 8).setValue(date); // Abriu_Em (se vazio)
                }
                destSheet.getRange(rowDest, 9).setValue(true); // Respondeu
            }

            atualizarContadoresProcesso(procAlvo.ID);
        });
    });

    return { novas_respostas: novas };
}

function baixarAnexo(msgId, filename) {
    try {
        const msg = GmailApp.getMessageById(msgId);
        const atts = msg.getAttachments();
        const target = atts.find(a => a.getName() === filename);
        if (!target) throw new Error('Anexo não encontrado');

        return {
            content: Utilities.base64Encode(target.getBytes()),
            contentType: target.getContentType(),
            name: target.getName()
        };
    } catch (e) {
        throw new Error('Erro ao baixar anexo: ' + e.toString());
    }
}

function todasRespostas() {
    return sheetToObjects(getSheet(ABA.respostas));
}

// ══════════════════════════════════════════════════════════════════
//  CONFIG
// ══════════════════════════════════════════════════════════════════
function lerTodasConfigs() {
    const sheet = getSheet(ABA.config);
    const rows = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
        : [];
    const cfg = {};
    rows.forEach(([k, v]) => { cfg[k] = v; });
    return cfg;
}

function salvarConfig(chave, valor) {
    const sheet = getSheet(ABA.config);
    const rows = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
        : [];
    for (let i = 0; i < rows.length; i++) {
        if (rows[i][0] === chave) {
            sheet.getRange(i + 2, 2).setValue(valor);
            return { ok: true };
        }
    }
    sheet.appendRow([chave, valor]);
    return { ok: true };
}

// ══════════════════════════════════════════════════════════════════
//  TRIGGERS — verificação automática
// ══════════════════════════════════════════════════════════════════
function ativarTrigger() {
    desativarTrigger(); // remove anteriores
    ScriptApp.newTrigger('verificarRespostas')
        .timeBased()
        .everyMinutes(15)
        .create();
    salvarConfig('verificacao_ativa', 'true');
    return { ok: true, msg: 'Verificação automática ativada (15 min)' };
}

function desativarTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
        if (t.getHandlerFunction() === 'verificarRespostas') ScriptApp.deleteTrigger(t);
    });
    salvarConfig('verificacao_ativa', 'false');
    return { ok: true, msg: 'Verificação automática desativada' };
}

function getMyEmailSafely() {
    try {
        return Session.getEffectiveUser().getEmail();
    } catch (e) {
        // Se falhar por permissão, tentamos pegar das configs da planilha
        const cfg = lerTodasConfigs();
        return cfg.gmailEmail || '';
    }
}
// ══════════════════════════════════════════════════════════════════
//  TEMPLATES
// ══════════════════════════════════════════════════════════════════
function listarTemplates() {
    const sheet = getOrCreateSheet(ABA.templates, ['ID', 'Nome', 'Categoria', 'Assunto', 'Corpo']);
    return sheetToObjects(sheet);
}

function salvarTemplate(d) {
    const sheet = getOrCreateSheet(ABA.templates, ['ID', 'Nome', 'Categoria', 'Assunto', 'Corpo']);
    const id = d.id || nextId(sheet);
    sheet.appendRow([id, d.name || '', d.cat || '', d.subject || '', d.body || '']);
    return { id };
}

function atualizarTemplate(d) {
    const sheet = getOrCreateSheet(ABA.templates, ['ID', 'Nome', 'Categoria', 'Assunto', 'Corpo']);
    const row = findRow(sheet, d.id);
    if (!row) throw new Error('Template não encontrado: ' + d.id);
    sheet.getRange(row, 1, 1, 5).setValues([[d.id, d.name || '', d.cat || '', d.subject || '', d.body || '']]);
    return { id: d.id };
}

function deletarTemplate(id) {
    const sheet = getOrCreateSheet(ABA.templates, ['ID', 'Nome', 'Categoria', 'Assunto', 'Corpo']);
    const row = findRow(sheet, id);
    if (row) sheet.deleteRow(row);
    return { ok: true };
}

// ══════════════════════════════════════════════════════════════════
//  AUTO-DISPATCH
// ══════════════════════════════════════════════════════════════════
function configurarAutoDispatch(d) {
    const sheet = getOrCreateSheet(ABA.autoDispatch, ['ProcessoID', 'IntervaloMs', 'MaxReenvios', 'ReenviosFeitos', 'ProximoEnvio', 'Ativo']);
    const now = new Date();
    const intervaloMs = Number(d.intervaloMs) || 0;
    const proximoEnvio = new Date(now.getTime() + intervaloMs).toISOString();

    // Atualiza se já existe, senão cria novo
    if (sheet.getLastRow() > 1) {
        const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
        const idx = rows.findIndex(r => String(r[0]) === String(d.processoId));
        if (idx !== -1) {
            sheet.getRange(idx + 2, 1, 1, 6).setValues([[
                d.processoId, intervaloMs, Number(d.maxReenvios) || 3, 0, proximoEnvio, true
            ]]);
            return { ok: true, updated: true };
        }
    }
    sheet.appendRow([d.processoId, intervaloMs, Number(d.maxReenvios) || 3, 0, proximoEnvio, true]);
    return { ok: true, created: true };
}

function processarAgendamentos() {
    const adSheet = getOrCreateSheet(ABA.autoDispatch, ['ProcessoID', 'IntervaloMs', 'MaxReenvios', 'ReenviosFeitos', 'ProximoEnvio', 'Ativo']);
    if (adSheet.getLastRow() < 2) return { processados: 0 };

    const now = new Date();
    const rows = adSheet.getRange(2, 1, adSheet.getLastRow() - 1, 6).getValues();
    let processados = 0;

    rows.forEach((row, i) => {
        const [processoId, intervaloMs, maxReenvios, reenviosFeitos, proximoEnvioStr, ativo] = row;
        const isAtivo = ativo === true || ativo === 'TRUE';
        if (!isAtivo) return;

        const proximoEnvio = new Date(proximoEnvioStr);
        if (isNaN(proximoEnvio) || proximoEnvio > now) return;

        // Busca dados do processo
        try {
            const procSheet = getSheet(ABA.processos);
            const destSheet = getSheet(ABA.destinatarios);
            const procRow = findRow(procSheet, processoId);
            if (!procRow) return;

            const procs = sheetToObjects(procSheet);
            const proc = procs.find(p => String(p.ID) === String(processoId));
            if (!proc) return;

            const dests = sheetToObjects(destSheet).filter(d => String(d.Processo_ID) === String(processoId));
            const novoReenvios = Number(reenviosFeitos) + 1;

            // Envia email para cada destinatário
            dests.forEach(dest => {
                try {
                    const corpo = (proc.Corpo_Email || '')
                        .replace(/{nome}/g, dest.Nome_Fornecedor || '')
                        .replace(/{empresa}/g, dest.Nome_Fornecedor || '')
                        .replace(/{tipo}/g, dest.Tipo_Material || '');
                    enviarEmail({
                        para: dest.Email_Fornecedor,
                        assunto: proc.Assunto,
                        corpo: corpo,
                        nome_remetente: lerTodasConfigs().gmail_nome || 'ShootMail',
                        processo_id: processoId,
                        fornecedor_id: dest.Fornecedor_ID,
                        nome_fornecedor: dest.Nome_Fornecedor,
                        tipo_material: dest.Tipo_Material,
                        nota: 'Auto-reenvio #' + novoReenvios
                    });
                } catch (e) {
                    console.error('Erro ao enviar auto-dispatch para ' + dest.Email_Fornecedor + ': ' + e);
                }
            });

            const rowIdx = i + 2;
            const novoProximoEnvio = new Date(proximoEnvio.getTime() + Number(intervaloMs)).toISOString();
            const novoAtivo = novoReenvios < Number(maxReenvios);
            adSheet.getRange(rowIdx, 4).setValue(novoReenvios);
            adSheet.getRange(rowIdx, 5).setValue(novoProximoEnvio);
            adSheet.getRange(rowIdx, 6).setValue(novoAtivo);
            processados++;
        } catch (e) {
            console.error('Erro no processamento do agendamento ' + processoId + ': ' + e);
        }
    });

    return { processados };
}

function ativarTriggerAutoDispatch() {
    // Remove triggers existentes para evitar duplicatas
    ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === 'processarAgendamentos') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('processarAgendamentos')
        .timeBased()
        .everyHours(1)
        .create();
    return { ok: true, msg: 'Trigger de auto-dispatch ativado (1 hora)' };
}

// ══════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════
function getSheet(nome) {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(nome);
    if (!sheet) throw new Error('Aba não encontrada: ' + nome + '. Execute setupPlanilha() primeiro.');
    return sheet;
}

function getOrCreateSheet(nome, headers) {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(nome);
    if (!sheet) {
        sheet = ss.insertSheet(nome);
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setBackground('#1c1f2e').setFontColor('#6b6f90').setFontWeight('bold');
        sheet.setFrozenRows(1);
    }
    return sheet;
}

function sheetToObjects(sheet) {
    if (sheet.getLastRow() < 2) return [];
    const [headers, ...rows] = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    return rows
        .filter(r => r[0] !== '' && r[0] !== null)
        .map(r => {
            const obj = {};
            headers.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i] : ''; });
            return obj;
        });
}

function nextId(sheet) {
    if (sheet.getLastRow() < 2) return 1;
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
        .map(v => parseInt(v) || 0);
    return Math.max(0, ...ids) + 1;
}

function findRow(sheet, id) {
    if (sheet.getLastRow() < 2) return null;
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const idx = ids.findIndex(v => String(v) === String(id));
    return idx === -1 ? null : idx + 2;
}

// ── RASTREAMENTO (PIXEL) ──
function handleTracking(e) {
    const pid = e.pid;
    const fid = e.fid;
    if (!pid || !fid) return returnPixel();

    try {
        const destSheet = getSheet(ABA.destinatarios);
        const row = findRowByMultiple(destSheet, [
            { col: 2, val: pid },
            { col: 3, val: fid }
        ]);

        if (row) {
            const alreadyOpened = destSheet.getRange(row, 7).getValue();
            if (!alreadyOpened || alreadyOpened === 'FALSE') {
                destSheet.getRange(row, 7).setValue(true);
                destSheet.getRange(row, 8).setValue(new Date().toISOString());
                atualizarContadoresProcesso(pid);
            }
        }
    } catch (err) {
        console.error('Erro no tracking:', err);
    }

    return returnPixel();
}

function returnPixel() {
    const clearGif = [
        0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00, 0xff, 0xff, 0xff,
        0x00, 0x00, 0x00, 0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00, 0x2c, 0x00, 0x00, 0x00, 0x00,
        0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02, 0x44, 0x01, 0x00, 0x3b
    ];
    return ContentService.createTextOutput(Utilities.newBlob(clearGif, 'image/gif').getDataAsString())
        .setMimeType(ContentService.MimeType.TEXT);
}

function findRowByMultiple(sheet, conditions) {
    if (sheet.getLastRow() < 2) return null;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const idx = data.findIndex(row =>
        conditions.every(c => String(row[c.col - 1]) === String(c.val))
    );
    return idx === -1 ? null : idx + 2;
}

function extractEmail(from) {
    const m = from.match(/<(.+?)>/);
    return m ? m[1] : from.trim();
}

function atualizarContadoresProcesso(procId) {
    const procSheet = getSheet(ABA.processos);
    const destSheet = getSheet(ABA.destinatarios);
    const respSheet = getSheet(ABA.respostas);

    const dests = sheetToObjects(destSheet).filter(d => String(d.Processo_ID) === String(procId));
    const resps = sheetToObjects(respSheet).filter(r => String(r.Processo_ID) === String(procId));

    const totalDisp = dests.reduce((a, d) => a + (parseInt(d.Total_Disparos_Individual) || 0), 0);
    const totalAbriu = dests.filter(d => d.Abriu === true || d.Abriu === 'TRUE').length;
    const totalResp = dests.filter(d => d.Respondeu === true || d.Respondeu === 'TRUE').length;
    const totalDest = dests.length;

    const procRow = findRow(procSheet, procId);
    if (!procRow) return;
    const status = totalResp > 0 ? 'replied' : 'active';
    procSheet.getRange(procRow, 5).setValue(status);
    procSheet.getRange(procRow, 6).setValue(totalDest);
    procSheet.getRange(procRow, 7).setValue(totalDisp);
    procSheet.getRange(procRow, 8).setValue(totalAbriu);
    procSheet.getRange(procRow, 9).setValue(totalResp);
    procSheet.getRange(procRow, 11).setValue(new Date().toISOString());
}
