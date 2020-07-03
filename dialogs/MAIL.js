/**
  __  __          _____ _      
 |  \/  |   /\   |_   _| |     
 | \  / |  /  \    | | | |     
 | |\/| | / /\ \   | | | |     
 | |  | |/ ____ \ _| |_| |____ 
 |_|  |_/_/    \_\_____|______|
                                                        
 */
const config = require('../config');
const nodeoutlook = require('nodejs-nodemailer-outlook');

const { ComponentDialog, WaterfallDialog, ChoicePrompt, TextPrompt,AttachmentPrompt, ChoiceFactory } = require('botbuilder-dialogs');

const MAIL_DIALOG = "MAIL_DIALOG";
const CHOICE_PROMPT = "CHOICE_PROMPT";
const TEXT_PROMPT = "TEXT_PROMPT";
const ATTACH_PROMPT = "ATTACH_PROMPT";
const WATERFALL_DIALOG = 'WATERFALL_DIALOG';

class MailDialog extends ComponentDialog {
    constructor(){
        super(MAIL_DIALOG);
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new AttachmentPrompt(ATTACH_PROMPT));
        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
          
            this.telefonoStep.bind(this),
            this.correoStep.bind(this),
            this.mailStep.bind(this)
        ]));
        this.initialDialogId = WATERFALL_DIALOG;
    }
 
    async telefonoStep(step){
        const details = step.options;
        console.log('[MailDialog]: telefonoDialog');
        return await step.prompt(TEXT_PROMPT, 'Escribe tu **teléfono y extención o celular**, para contactarte.');
    }
    async correoStep(step){
        console.log('[MailDialog]: correoStep');
        const details = step.options;

        const tel = step.result;
        details.tel = tel;
    
        console.log(details.tel);
        return await step.prompt(TEXT_PROMPT, 'Escribe tu **correo electrónico** para enviarte los detalles del servicio.');
    }
    async mailStep(step){
        console.log('[MailDialog]: mailStep');
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
                subject: `${details.proyecto} Error datos de resguardo:  ${details.serie}`,
                html: `<p>El usuario <b>${details.nombre}</b>, ha indicado que hay un error en la información de su resguardo con el número de serie ${details.serie}:</p>
                <hr>
                
                
                <p><b>Datos de contacto: </b></p>
                <p>Teléfono: ${details.tel}</p>
                <p>Email: ${details.email}</p>
                

                <hr>
                <p>Datos del equipo reportado:</p><br> 
                <b>Proyecto: ${details.proyecto}</b>  <br> 
                <b>Serie: ${details.serie}</b> <br> 
                <b>Perfil: ${details.perfil}</b> <br> 
                <b>Usuario: ${details.nombre}</b> <br> 
                <b>Estado: ${details.estado}</b> <br> 
                <b>Municipio: ${details.municipio}</b> <br> 
                <b>Dirección: ${details.direccion}</b> <br> 
                <b>Teléfono: ${details.telefono}</b> <br> 
                
                <hr> 
                <p>En breve nuestro ingeniero se comunicará con usted.</p>
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

        await step.context.sendActivity(`**Gracias por tu apoyo, se enviará un correo con tus datos**`);
        await step.cancelAllDialogs();
        return await step.endDialog();

    }   

}
module.exports.MailDialog = MailDialog;
module.exports.MAIL_DIALOG = MAIL_DIALOG;