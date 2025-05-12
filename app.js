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
        status: d.STATUS || 'NÃO ENVIADO'
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
                    <p>Olá! Este e-mail foi enviado automaticamente.</p>
                    <img src="http://localhost:3000/pixel?email=${encodeURIComponent(contato.email)}" width="1" height="1" style="display:none;">
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


app.get('/pixel', (req, res) => {
    const email = req.query.email;
    if (!email) return res.status(400).send('Email não informado.');

    try {
        const workbook = xlsx.readFile('./dados.xlsx');
        const sheet = workbook.Sheets['Planilha1'];
        const data = xlsx.utils.sheet_to_json(sheet, { defval: '' });

        // Procura a linha com o email correspondente
        const linhaIndex = data.findIndex(row => row.EMAIL?.toLowerCase() === email.toLowerCase());
        if (linhaIndex >= 0) {
            const visualizadoColIndex = Object.keys(data[0]).findIndex(col => col.toUpperCase() === 'VISUALIZADO');
            if (visualizadoColIndex === -1) return res.status(500).send('Coluna VISUALIZADO não encontrada');

            const colLetra = String.fromCharCode(65 + visualizadoColIndex); // A, B, C...
            const excelRow = linhaIndex + 2; // +2 porque JSON começa do índice 0 e Excel da linha 2 (1-based + cabeçalho)
            const celula = `${colLetra}${excelRow}`;

            sheet[celula] = { t: 's', v: 'SIM' };
            xlsx.writeFile(workbook, './dados.xlsx');
        }

        // Retorna o pixel transparente
        const pixelGif = Buffer.from('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==', 'base64');
        res.setHeader('Content-Type', 'image/gif');
        res.end(pixelGif);
    } catch (err) {
        console.error("Erro no pixel:", err);
        res.status(500).send('Erro interno.');
    }
});



app.listen(PORT, () => {
    console.log(`🚀 Servidor rodando em: http://localhost:${PORT}`);
});
