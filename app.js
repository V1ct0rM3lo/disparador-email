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

 return data.filter(d => d.SITUACAO === 'A' && d.EMAIL).map(d => ({
    cod: d.COD_EMPRESA,
    nome: d.NOME_EMPRESA,
    cnpj: d.CNPJ,
    email: d.EMAIL,
    situacao: d.SITUACAO,
    status: d.STATUS || 'NÃO ENVIADO',
    visualizado: d.VISUALIZADO || ''
}));

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

    // Localiza colunas
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = sheet[xlsx.utils.encode_cell({ r: 0, c: C })];
        if (cell && cell.v) {
            const header = cell.v.toString().toUpperCase();
            if (header === 'STATUS') statusCol = C;
            if (header === 'EMAIL') emailCol = C;
            if (header === 'COD_EMPRESA') codCol = C;
        }
    }

    if (statusCol === null || emailCol === null || codCol === null) {
        return res.status(500).send({ error: 'Colunas STATUS, EMAIL ou COD_EMPRESA não encontradas.' });
    }

    for (const contato of selecionados) {
        try {
            await transporter.sendMail({
                from: `"Disparador" <${process.env.EMAIL_USER}>`,
                to: contato.email,
                subject: '🔔 Notificação do Sistema - Disparo Automático',
                html: `
    <div style="font-family: Arial, sans-serif; background-color: #f4f4f4; padding: 20px;">
      <div style="background-color: #1e1e2f; color: #ffffff; padding: 15px; border-radius: 8px 8px 0 0;">
        <h2 style="margin: 0;">🚀 Sistema de Disparo Automático</h2>
      </div>

      <div style="background-color: #ffffff; padding: 20px; border-radius: 0 0 8px 8px;">
        <p>Olá, tudo certo? 🤖</p>

        <p>Este e-mail foi enviado automaticamente pelo nosso sistema Node.js como parte de um teste de funcionalidade.</p>

        <p>Você pode acessar nosso painel clicando no botão abaixo:</p>

        <a href="https://seusite.com/painel" 
           style="display: inline-block; padding: 12px 24px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px; font-weight: bold;">
          🔗 Acessar Painel
        </a>

        <p style="margin-top: 20px; font-size: 12px; color: #888;">Se você não solicitou este e-mail, apenas ignore esta mensagem.</p>

      <img src="https://disparador-email.onrender.com/pixel?email=${encodeURIComponent(contato.email)}" width="1" height="1" style="display:none;">

      </div>
    </div>
    `
            });

            console.log(`✅ E-mail enviado para ${contato.email}`);

            // Atualiza status na planilha
            for (let R = 1; R <= range.e.r; ++R) {
                const emailCell = xlsx.utils.encode_cell({ r: R, c: emailCol });
                const codCell = xlsx.utils.encode_cell({ r: R, c: codCol });
                const statusCell = xlsx.utils.encode_cell({ r: R, c: statusCol });

                const emailVal = sheet[emailCell]?.v?.toString().trim().toLowerCase();
                const codVal = sheet[codCell]?.v?.toString().trim();
                const statusVal = sheet[statusCell]?.v?.toString().trim().toUpperCase();

                if (
                    emailVal === contato.email.toLowerCase().trim() &&
                    codVal === contato.cod.toString() &&
                    statusVal === 'NÃO ENVIADO'
                ) {
                    sheet[statusCell] = { t: 's', v: 'ENVIADO' };
                    console.log(`📌 STATUS atualizado na linha ${R + 1} (${statusCell})`);
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
      const novaSheet = xlsx.utils.json_to_sheet(dados);
      workbook.Sheets[sheetName] = novaSheet;
      xlsx.writeFile(workbook, './dados.xlsx');
    } else {
      console.log(`⚠️ Nenhuma correspondência encontrada para ${email}`);
    }
  }

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
});




app.listen(PORT, () => {
    console.log(`🚀 Servidor rodando em: http://localhost:${PORT}`);
});
