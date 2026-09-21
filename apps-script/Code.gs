const RIFAS = {
  'RIFA-0003': {
    aba: 'Controle RIFA-0003',
    valorNumero: 10,
    descricao: 'RIFA-0003 — Chuteira Umbro ou R$ 200 em compras na Escolástico Esportes',
    exigeTime: false
  },
  'RIFA-0001': {
    aba: 'Controle',
    valorNumero: 10,
    descricao: 'Rifa do Dia das Crianças',
    exigeTime: false
  },
  'RIFA-0002': {
    aba: 'Controle RIFA-0002',
    valorNumero: 5,
    descricao: 'RIFA-0002 — Camisa do seu time',
    exigeTime: true
  }
};

function respostaJSON(conteudo) {
  return ContentService
    .createTextOutput(JSON.stringify(conteudo))
    .setMimeType(ContentService.MimeType.JSON);
}

function obterConfigRifa(idInformado) {
  const id = String(idInformado || 'RIFA-0001').trim().toUpperCase();
  const config = RIFAS[id];

  if (!config) {
    throw new Error('Rifa não encontrada.');
  }

  return {
    id: id,
    aba: config.aba,
    valorNumero: config.valorNumero,
    descricao: config.descricao,
    exigeTime: config.exigeTime
  };
}

function obterAba(planilha, config) {
  const aba = planilha.getSheetByName(config.aba);

  if (!aba) {
    throw new Error('A aba "' + config.aba + '" não foi encontrada.');
  }

  return aba;
}

function doGet(e) {
  try {
    const config = obterConfigRifa(
      e && e.parameter ? e.parameter.rifa : 'RIFA-0001'
    );
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = obterAba(planilha, config);
    const dados = aba.getDataRange().getDisplayValues();

    if (dados.length < 2) {
      return respostaJSON({
        sucesso: true,
        rifa: config.id,
        total: 0,
        numeros: []
      });
    }

    const cabecalhos = dados[0].map(function(cabecalho) {
      return String(cabecalho).trim();
    });

    const numeros = dados.slice(1)
      .filter(function(linha) {
        return String(linha[0] || '').trim() !== '';
      })
      .map(function(linha) {
        const registro = {};

        cabecalhos.forEach(function(cabecalho, indice) {
          if (cabecalho !== '') {
            registro[cabecalho] = linha[indice] || '';
          }
        });

        return registro;
      });

    return respostaJSON({
      sucesso: true,
      rifa: config.id,
      total: numeros.length,
      atualizadoEm: new Date().toISOString(),
      numeros: numeros
    });
  } catch (erro) {
    return respostaJSON({
      sucesso: false,
      erro: erro.message
    });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      return respostaJSON({
        sucesso: false,
        erro: 'Nenhum dado foi recebido.'
      });
    }

    const pedido = JSON.parse(e.postData.contents);
    const config = obterConfigRifa(pedido.rifa);
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const aba = obterAba(planilha, config);
    const nome = String(pedido.nome || '').trim();
    const time = String(pedido.time || '').trim();
    const whatsappInformado = String(pedido.whatsapp || '').replace(/\D/g, '');
    let whatsapp = whatsappInformado;

    if (
      (whatsapp.length === 12 || whatsapp.length === 13) &&
      whatsapp.indexOf('55') === 0
    ) {
      whatsapp = whatsapp.substring(2);
    }

    const numerosSolicitados = Array.isArray(pedido.numeros)
      ? pedido.numeros.map(function(numero) {
          return String(numero).trim();
        })
      : [];

    if (nome.length < 2) {
      return respostaJSON({
        sucesso: false,
        erro: 'Informe o nome do comprador.'
      });
    }

    if (whatsapp.length < 10 || whatsapp.length > 11) {
      return respostaJSON({
        sucesso: false,
        erro: 'Informe um WhatsApp válido com DDD.'
      });
    }

    if (config.exigeTime && time === '') {
      return respostaJSON({
        sucesso: false,
        erro: 'Informe o time escolhido.'
      });
    }

    if (numerosSolicitados.length === 0) {
      return respostaJSON({
        sucesso: false,
        erro: 'Nenhum número foi selecionado.'
      });
    }

    const numerosUnicos = [...new Set(numerosSolicitados)];
    const ultimaColuna = aba.getLastColumn();
    const cabecalhos = aba
      .getRange(1, 1, 1, ultimaColuna)
      .getDisplayValues()[0]
      .map(function(valor) {
        return String(valor).trim();
      });

    function coluna(nomeCabecalho) {
      return cabecalhos.indexOf(nomeCabecalho);
    }

    const colNumero = coluna('Número');
    const colStatus = coluna('Status');
    const colNome = coluna('Nome do comprador');
    const colWhatsApp = coluna('WhatsApp');
    const colValor = coluna('Valor (R$)');
    const colForma = coluna('Forma de pagamento');
    const colData = coluna('Data do pagamento');
    const colObs = coluna('Observações');
    const colContatoWhatsApp = coluna('Contato WhatsApp');
    const colDataReserva = coluna('Data/Hora da Reserva');
    const colContatoRealizado = coluna('Contato realizado');
    const colTime = coluna('Time escolhido');

    if (colNumero === -1 || colStatus === -1) {
      return respostaJSON({
        sucesso: false,
        erro: 'As colunas Número e Status não foram encontradas.'
      });
    }

    if (
      colContatoWhatsApp === -1 ||
      colDataReserva === -1 ||
      colContatoRealizado === -1
    ) {
      return respostaJSON({
        sucesso: false,
        erro: 'As colunas administrativas da reserva não foram encontradas.'
      });
    }

    if (config.exigeTime && colTime === -1) {
      return respostaJSON({
        sucesso: false,
        erro: 'A coluna Time escolhido não foi encontrada.'
      });
    }

    const ultimaLinha = aba.getLastRow();

    if (ultimaLinha < 2) {
      return respostaJSON({
        sucesso: false,
        erro: 'Não existem números cadastrados.'
      });
    }

    const valores = aba
      .getRange(2, 1, ultimaLinha - 1, ultimaColuna)
      .getValues();
    const linhasParaReservar = [];
    const indisponiveis = [];
    const inexistentes = [];

    numerosUnicos.forEach(function(numeroSolicitado) {
      let encontrado = false;

      for (let i = 0; i < valores.length; i++) {
        const numeroPlanilha = String(valores[i][colNumero]).trim();

        if (Number(numeroPlanilha) === Number(numeroSolicitado)) {
          encontrado = true;
          const status = String(valores[i][colStatus] || '')
            .trim()
            .toLowerCase();

          if (
            status === '' ||
            status === 'disponível' ||
            status === 'disponivel'
          ) {
            linhasParaReservar.push(i + 2);
          } else {
            indisponiveis.push(numeroSolicitado);
          }

          break;
        }
      }

      if (!encontrado) {
        inexistentes.push(numeroSolicitado);
      }
    });

    if (indisponiveis.length > 0) {
      return respostaJSON({
        sucesso: false,
        conflito: true,
        erro: 'Um ou mais números já foram reservados ou pagos.',
        numerosIndisponiveis: indisponiveis
      });
    }

    if (inexistentes.length > 0) {
      return respostaJSON({
        sucesso: false,
        erro: 'Um ou mais números não foram encontrados.',
        numerosInexistentes: inexistentes
      });
    }

    const agora = new Date();
    const whatsappComPais = '55' + whatsapp;
    const numerosFormatados = numerosUnicos.map(function(numero) {
      return String(Number(numero)).padStart(2, '0');
    });
    const complementoTime = config.exigeTime
      ? ' Você escolheu a camisa do ' + time + '.'
      : '';
    const mensagemWhatsApp =
      'Olá, ' + nome + '! Aqui é da Associação Atlética Parque São Luiz – PSL. ' +
      'Vi que você reservou o(s) número(s) ' + numerosFormatados.join(', ') +
      ' na ' + config.descricao + '.' + complementoTime +
      ' Estou entrando em contato para saber se conseguiu concluir o pagamento via Pix. ' +
      'Se precisar de ajuda, estou à disposição.';
    const linkWhatsApp =
      'https://wa.me/' + whatsappComPais +
      '?text=' + encodeURIComponent(mensagemWhatsApp);
    const textoComLink = SpreadsheetApp.newRichTextValue()
      .setText('Abrir WhatsApp')
      .setLinkUrl(linkWhatsApp)
      .build();

    linhasParaReservar.forEach(function(linha) {
      aba.getRange(linha, colStatus + 1).setValue('Reservado');

      if (colNome !== -1) {
        aba.getRange(linha, colNome + 1).setValue(nome);
      }

      if (colWhatsApp !== -1) {
        aba.getRange(linha, colWhatsApp + 1).setValue(whatsapp);
      }

      if (colValor !== -1) {
        aba.getRange(linha, colValor + 1).setValue(config.valorNumero);
      }

      if (colForma !== -1) {
        aba.getRange(linha, colForma + 1).setValue('Pix');
      }

      if (colData !== -1) {
        aba.getRange(linha, colData + 1).setValue(agora);
      }

      if (colObs !== -1) {
        aba.getRange(linha, colObs + 1)
          .setValue('Reserva realizada pelo site — ' + config.id);
      }

      aba.getRange(linha, colContatoWhatsApp + 1)
        .setRichTextValue(textoComLink);
      aba.getRange(linha, colDataReserva + 1).setValue(agora);
      aba.getRange(linha, colContatoRealizado + 1).setValue('Não');

      if (config.exigeTime && colTime !== -1) {
        aba.getRange(linha, colTime + 1).setValue(time);
      }
    });

    SpreadsheetApp.flush();

    return respostaJSON({
      sucesso: true,
      rifa: config.id,
      mensagem: 'Reserva realizada com sucesso.',
      nome: nome,
      whatsapp: whatsapp,
      time: time,
      numeros: numerosUnicos,
      quantidade: numerosUnicos.length,
      valorTotal: numerosUnicos.length * config.valorNumero
    });
  } catch (erro) {
    return respostaJSON({
      sucesso: false,
      erro: erro.message
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (erroLock) {}
  }
}
