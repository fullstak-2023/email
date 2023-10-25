const fs = require('fs');
const { google } = require('googleapis');
const path = require('path');
const process = require('process');
const { authenticate } = require('@google-cloud/local-auth');
const docx = require("docx");

const SCOPES = ['https://www.googleapis.com/auth/gmail.send'];
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

async function loadSavedCredentialsIfExist() {
    try {
        const content = await fs.promises.readFile(TOKEN_PATH);
        const credentials = JSON.parse(content);
        return google.auth.fromJSON(credentials);
    } catch (err) {
        return null;
    }
}

async function saveCredentials(client) {
    const content = await fs.promises.readFile(CREDENTIALS_PATH);
    const keys = JSON.parse(content);
    const key = keys.installed || keys.web;
    const payload = JSON.stringify({
        type: 'authorized_user',
        client_id: key.client_id,
        client_secret: key.client_secret,
        refresh_token: client.credentials.refresh_token,
    });
    await fs.promises.writeFile(TOKEN_PATH, payload);
}

async function authorize() {
    let client = await loadSavedCredentialsIfExist();
    if (client) {
        return client;
    }
    client = await authenticate({
        scopes: SCOPES,
        keyfilePath: CREDENTIALS_PATH,
    });
    if (client.credentials) {
        await saveCredentials(client);
    }
    return client;
}

async function sendEmail(auth, toEmailAddress, subject, message, attachedFileBuffer) {
    const gmail = google.gmail({ version: 'v1', auth });
    const fileData = attachedFileBuffer.toString('base64');
    const rawEmail = makeEmail(toEmailAddress, subject, message, fileData);
    const res = await gmail.users.messages.send({
        userId: 'me',
        resource: {
            raw: rawEmail,
            labelIds: ['INBOX'],
        },
    });
    console.log('החלשנ העדוהה');
    // console.log('Email sent:', res.data);
}

function myWord(myName) {
    return new Promise((resolve, reject) => {
        const doc = new docx.Document({
            sections: [
                {
                    properties: {},
                    children: [
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: `Happy New Year to dear ${myName}`,
                                    size: 100,
                                    bold: true,
                                }),
                            ],
                        }),
                    ],
                },
            ],
        });
        docx.Packer.toBuffer(doc)
            .then((buffer) => {
                fs.writeFileSync(`${myName}.docx`, buffer);
                console.log("החלצהב רצונ ץבוקה");
                resolve();
            })
            .catch((error) => {
                reject(error);
            });
    });
}

async function getDocx(myName) {
    const DocxBuffer = await fs.promises.readFile(`${myName}.docx`);
    return DocxBuffer;
}

function makeEmail(to, subject, message, fileData) {
    const email = [
        'Content-Type: text/plain; charset="UTF-8"\r\n',
        'MIME-Version: 1.0\r\n',
        `To: ${to}\r\n`,
        `Subject: ${subject}\r\n`,
        `Content-Disposition: attachment; filename="shana_tova.docx"\r\n`,
        'Content-Transfer-Encoding: base64\r\n\r\n',
        message,
        '\r\n\r\n',
        fileData,
    ].join('');

    return Buffer.from(email).toString('base64');
}

class ShanaTova {
    async getNameAndEmail(myName, myEmail) {
        await myWord(myName);
        const DocxBuffer = await getDocx(myName);

        try {
            const auth = await authorize();
            const toEmailAddress = myEmail;
            const subject = "Hello " + myName;
            const message = 'אני מאחל לך שנה טובה ומתוקה';
            await sendEmail(auth, toEmailAddress, subject, message, DocxBuffer);
        } catch (error) {
            console.error(error);
        }
    }
}

module.exports = new ShanaTova();
