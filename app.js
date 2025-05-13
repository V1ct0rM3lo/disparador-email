require('dotenv').config();
const express = require('express');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const bodyParser = require('body-parser');
const app = express();
const PORT = 3000;

app.use(express.static('public'));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Página principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'emails.html'));
});

// Transportador de e-mail
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
    }
});

// Lê os contatos da planilha
function getContatosAtivos() {
    const workbook = xlsx.readFile('./dados.xlsx');
    const sheet = workbook.Sheets['Planilha1'];
    const data = xlsx.utils.sheet_to_json(sheet, { defval: '' });

    let atualizou = false;

    const contatos = data.filter(d => d.SITUACAO === 'A' && d.EMAIL).map(d => {
        // Preenche STATUS vazio
        if (!d.STATUS || d.STATUS.toString().trim() === '') {
            d.STATUS = 'NÃO ENVIADO';
            atualizou = true;
        }

        // Preenche VISUALIZADO vazio
        if (!d.VISUALIZADO || d.VISUALIZADO.toString().trim() === '') {
            d.VISUALIZADO = 'NÃO VISUALIZADO';
            atualizou = true;
        }

        return {
            cod: d.COD_EMPRESA,
            nome: d.NOME_EMPRESA,
            cnpj: d.CNPJ,
            email: d.EMAIL,
            situacao: d.SITUACAO,
            status: d.STATUS,
            visualizado: d.VISUALIZADO
        };
    });

    if (atualizou) {
        const novaSheet = xlsx.utils.json_to_sheet(data);
        workbook.Sheets['Planilha1'] = novaSheet;
        xlsx.writeFile(workbook, './dados.xlsx');
        console.log('📁 Planilha atualizada: STATUS e VISUALIZADO preenchidos automaticamente onde estavam vazios.');
    }

    return contatos;
}

// Retorna contatos ativos
app.get('/contatos', (req, res) => {
    const contatos = getContatosAtivos();
    res.json(contatos);
});

// Envia e-mails e atualiza STATUS
app.post('/enviar-emails', async (req, res) => {
    const selecionados = req.body.emails;

    const workbook = xlsx.readFile('./dados.xlsx');
    const sheet = workbook.Sheets['Planilha1'];
    const data = xlsx.utils.sheet_to_json(sheet, { defval: '' });
    const range = xlsx.utils.decode_range(sheet['!ref']);

    let statusCol = null, emailCol = null, codCol = null;
let visualizadoCol = null;

// Localiza também a coluna VISUALIZADO
for (let C = range.s.c; C <= range.e.c; ++C) {
    const cell = sheet[xlsx.utils.encode_cell({ r: 0, c: C })];
    if (cell && cell.v) {
        const header = cell.v.toString().toUpperCase();
        if (header === 'STATUS') statusCol = C;
        if (header === 'EMAIL') emailCol = C;
        if (header === 'COD_EMPRESA') codCol = C;
        if (header === 'VISUALIZADO') visualizadoCol = C;
    }
}

// Verifica se todas foram encontradas
if (statusCol === null || emailCol === null || codCol === null || visualizadoCol === null) {
    return res.status(500).send({ error: 'Colunas obrigatórias não encontradas.' });
}

// Atualiza STATUS e VISUALIZADO
for (const contato of selecionados) {
    try {
        await transporter.sendMail({
            from: `"Disparador" <${process.env.EMAIL_USER}>`,
            to: contato.email,
            subject: '🔔 Notificação do Sistema - Disparo Automático',
            html: `...` // sua estrutura de e-mail permanece
        });

        console.log(`✅ E-mail enviado para ${contato.email}`);

        for (let R = 1; R <= range.e.r; ++R) {
            const emailCell = xlsx.utils.encode_cell({ r: R, c: emailCol });
            const codCell = xlsx.utils.encode_cell({ r: R, c: codCol });
            const statusCell = xlsx.utils.encode_cell({ r: R, c: statusCol });
            const visualizadoCell = xlsx.utils.encode_cell({ r: R, c: visualizadoCol });

            const emailVal = sheet[emailCell]?.v?.toString().trim().toLowerCase();
            const codVal = sheet[codCell]?.v?.toString().trim();
            const statusVal = sheet[statusCell]?.v?.toString().trim().toUpperCase();

            if (
                emailVal === contato.email.toLowerCase().trim() &&
                codVal === contato.cod.toString()
            ) {
                sheet[statusCell] = { t: 's', v: 'ENVIADO' };
                sheet[visualizadoCell] = { t: 's', v: 'NÃO VISUALIZADO' };
                console.log(`📌 STATUS e VISUALIZADO atualizados na linha ${R + 1}`);
                break;
            }
        }

    } catch (err) {
        console.error(`❌ Erro ao enviar e-mail: ${err.message}`);
    }
}


    xlsx.writeFile(workbook, './dados.xlsx');
    console.log('✅ Planilha atualizada com sucesso.');
    res.send({ status: "ok", enviados: selecionados.length });
});


app.get('/pixel', async (req, res) => {
  const email = req.query.email;

  if (email) {
    const workbook = xlsx.readFile('./dados.xlsx');
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const dados = xlsx.utils.sheet_to_json(sheet, { defval: '' });

    let atualizado = false;

    for (let i = 0; i < dados.length; i++) {
      if (dados[i].EMAIL && dados[i].EMAIL.trim().toLowerCase() === email.trim().toLowerCase()) {
        dados[i].VISUALIZADO = `Visualização registrada para ${email}`;
        atualizado = true;
        console.log(`👀 Visualização registrada para ${email}`);
        break;
      }
    }

    if (atualizado) {
   const range = xlsx.utils.decode_range(sheet['!ref']);
let emailCol = null;
let visualizadoCol = null;

// Acha os índices das colunas
for (let C = range.s.c; C <= range.e.c; ++C) {
  const header = sheet[xlsx.utils.encode_cell({ r: 0, c: C })];
  if (header && header.v) {
    const colName = header.v.toString().toUpperCase();
    if (colName === 'EMAIL') emailCol = C;
    if (colName === 'VISUALIZADO') visualizadoCol = C;
  }
}

if (emailCol !== null && visualizadoCol !== null) {
  for (let R = 1; R <= range.e.r; ++R) {
    const emailCell = xlsx.utils.encode_cell({ r: R, c: emailCol });
    const visualizadoCell = xlsx.utils.encode_cell({ r: R, c: visualizadoCol });

    const emailVal = sheet[emailCell]?.v?.toString().trim().toLowerCase();

    if (emailVal === email.trim().toLowerCase()) {
      sheet[visualizadoCell] = {
        t: 's',
        v: `Visualização registrada para ${email}`
      };
      console.log(`👀 Visualização registrada diretamente na célula ${visualizadoCell}`);
      break;
    }
  }

  xlsx.writeFile(workbook, './dados.xlsx');
}


  // Resposta para o pixel
     // Resposta para o pixel
    const imgBuffer = Buffer.from(
      "R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==",
      "base64"
    );

    res.writeHead(200, {
      'Content-Type': 'image/gif',
      'Content-Length': imgBuffer.length
    });
    res.end(imgBuffer);
  } // fecha o if (email)
} // fecha o app.get('/pixel', ...)
);

// Inicia o servidor
app.listen(PORT, () => {
  console.log(`🚀 Servidor rodando em: http://localhost:${PORT}`);
});


