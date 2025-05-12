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

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'emails.html'));
});


// Transportador de e-mail
const transporter = nodemailer.createTransport({
    service: 'gmail', // ou outro
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
    }
});

// Lê dados da planilha
// Lê dados da planilha
// Lê dados da planilha
function getContatosAtivos() {
    const workbook = xlsx.readFile('./dados.xlsx');
    const sheet = workbook.Sheets['Planilha1'];
    const data = xlsx.utils.sheet_to_json(sheet);

    return data.filter(d => d.SITUACAO === 'A' && d.EMAIL).map(d => ({
        cod: d.COD_EMPRESA,         // Código da empresa
        nome: d.NOME_EMPRESA,       // Nome da empresa
        cnpj: d.CNPJ,               // CNPJ
        email: d.EMAIL,             // Email
        situacao: d.SITUACAO        // Situação
    }));
}

// Rota para fornecer os dados
app.get('/contatos', (req, res) => {
    const contatos = getContatosAtivos();
    res.json(contatos);
});

// Rota para envio dos e-mails
app.post('/enviar-emails', async (req, res) => {
    const selecionados = req.body.emails;

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

    <img src="http://localhost:3000/pixel?email=${encodeURIComponent(contato.email)}" width="1" height="1" style="display:none;">
  </div>
</div>
`

            });

            console.log(`✅ Enviado para ${contato.email}`);
        } catch (err) {
            console.error(`❌ Erro com ${contato.email}: ${err.message}`);
        }
    }

    res.send({ status: "ok", enviados: selecionados.length });
});

app.listen(PORT, () => {
    console.log(`Servidor no ar: http://localhost:${PORT}`);
});


app.post('/atualizar-status/:codEmpresa', async (req, res) => {
    const codEmpresa = req.params.codEmpresa;
    const { status } = req.body;

    console.log(`Status da empresa ${codEmpresa} atualizado para: ${status}`);

    // Apenas envia uma resposta simulando sucesso
    res.status(200).json({ message: 'Status atualizado (simulado).' });
});

app.get('/pixel', (req, res) => {
    const email = req.query.email;
    const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress;

    console.log(`E-mail aberto por: ${email} - IP: ${ip} - ${new Date().toISOString()}`);

    const img = Buffer.from(
        'R0lGODlhAQABAPAAAAAAAAAAACH5BAAAAAAALAAAAAABAAEAAAICRAEAOw==',
        'base64'
    );

    res.writeHead(200, {
        'Content-Type': 'image/gif',
        'Content-Length': img.length
    });

    res.end(img);
});
