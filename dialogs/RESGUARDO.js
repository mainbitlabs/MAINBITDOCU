/**                                         
 ____  _____ ____   ____ _   _   _    ____  ____   ___  
|  _ \| ____/ ___| / ___| | | | / \  |  _ \|  _ \ / _ \ 
| |_) |  _| \___ \| |  _| | | |/ _ \ | |_) | | | | | | |
|  _ <| |___ ___) | |_| | |_| / ___ \|  _ <| |_| | |_| |
|_| \_|_____|____/ \____|\___/_/   \_|_| \_|____/ \___/ 
                                                                                                        
 */
const config = require('../config');
const nodeoutlook = require('nodejs-nodemailer-outlook');

const { ComponentDialog, WaterfallDialog, ChoicePrompt, TextPrompt,AttachmentPrompt, ChoiceFactory } = require('botbuilder-dialogs');

const RESGUARDO_DIALOG = "RESGUARDO_DIALOG";
const CHOICE_PROMPT = "CHOICE_PROMPT";
const TEXT_PROMPT = "TEXT_PROMPT";
const ATTACH_PROMPT = "ATTACH_PROMPT";
const WATERFALL_DIALOG = 'WATERFALL_DIALOG';

class ResguardoDialog extends ComponentDialog {
    constructor(){
        super(RESGUARDO_DIALOG);
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new AttachmentPrompt(ATTACH_PROMPT));
        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
          
            this.correoStep.bind(this),
            this.mailStep.bind(this)
        ]));
        this.initialDialogId = WATERFALL_DIALOG;
    }
 
  
    async correoStep(step){
        console.log('[ResguardoDialog]: correoStep');
        const details = step.options;

        const tel = step.result;
        details.tel = tel;
    
        console.log(details.tel);
        return await step.prompt(TEXT_PROMPT, 'Escribe el **correo electrónico** donde será enviado tu resguardo.');
    }
    async mailStep(step){
        console.log('[ResguardoDialog]: mailStep');
        const details = step.options;
        const mail = step.result;
        details.email = mail;
        console.log(details.email);

        const email = new Promise((resolve, reject) => { 
            nodeoutlook.sendEmail({
                auth: {
                    user: `${config.email1}`,
                    pass: `${config.pass}`,
                }, from: `${config.email1}`,
                to: `${config.email3}`,
                bcc: `${details.email}`,
                subject: `${details.proyecto} Formato de resguardo:  ${details.serie}`,
                html: `<p>Buen día <b>${details.nombre}</b>, te hacemos llegar el formato de resguardo para el equipo de cómputo con el número de serie <b>${details.serie}</b>:</p>
                <hr>
                
                
                <p><b>Debes seguir las siguientes indicaciones: </b></p>
                <ol>
                    <li><b>Imprime</b> el formato adjunto</li>
                    <li><b>Revisa</b> que todos los datos sean correctos</li>
                    <li><b>Firma</b> en la parte correspondiente</li>
                    <li><b>Envía</b> el formato firmado a través del Bot <b>MAINBITDOCU</b></li>
                </ol>
                
                <hr> 
                <p>Un placer atenderle.</p>
                <p>Equipo Mainbit.</p>
                <hr>
                <img style="width:100%" src="https://raw.githubusercontent.com/esanchezlMBT/images/master/firma2020.jpg">
                `,
                onError: (e) => reject(console.log(e)),
                onSuccess: (i) => resolve(console.log(i))
                }
            );
            
        });

        await email;

        await step.context.sendActivity(`**Gracias por su apoyo, en breve te será enviado un correo con el formato de resguardo para que lo firmes.**`);
        await step.cancelAllDialogs();
        return await step.endDialog();

    }   

}
module.exports.ResguardoDialog = ResguardoDialog;
module.exports.RESGUARDO_DIALOG = RESGUARDO_DIALOG;