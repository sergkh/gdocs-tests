import XmlDocument = GoogleAppsScript.XML_Service.Document;
import { Question } from './Question';

export class MoodleExporter {
    export(questions: Question[]): XmlDocument {
        const el = (name) => XmlService.createElement(name);
        const cdata = (el, text) => el.addContent(XmlService.createCdata(text));

        var doc = XmlService.createDocument();
        var quizes = el('quiz');
        doc.addContent(quizes);

        questions.forEach((q) => {
            if (q.options.length <= 1) {
                Logger.log('Warning skipping freeform question: ' + q.text);
                return;
            }

            const question = el('question').setAttribute('type', 'multichoice');

            question.addContent(el('name').addContent(cdata(el('text'), q.id)));

            question.addContent(
                el('questiontext').addContent(cdata(el('text'), q.textAsHtml()))
            );

            question.addContent(el('hidden').setText('0'));
            question.addContent(el('penalty').setText('0.3333333'));
            question.addContent(el('idnumber').setText(q.id + ''));
            question.addContent(
                el('single').setText(
                    q.correctOptions.length > 1 ? 'false' : 'true'
                )
            );
            question.addContent(el('defaultgrade').setText('2.000'));
            question.addContent(el('shuffleanswers').setText('true'));
            question.addContent(el('answernumbering').setText('abc'));
            question.addContent(el('showstandardinstruction').setText('1'));

            q.options.forEach((answer, idx) => {
                const correct = q.correctOptions.includes(idx + 1);
                const node = el('answer')
                    .setAttribute(
                        'fraction',
                        correct ? '' + 100 / q.correctOptions.length : '0'
                    )
                    .setAttribute('format', 'html')
                    .addContent(cdata(el('text'), answer));

                question.addContent(node);
            });

            const tags = el('tags');
            q.tags.forEach((t) =>
                tags.addContent(el('tag').addContent(el('text').setText(t + '')))
            );
            question.addContent(tags);

            quizes.addContent(question);
        });

        return doc;
    }
}
