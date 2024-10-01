import { GenerateDocumentService } from "./GenerateDocumentService";
import { MoodleExporter } from "./MoodleExporter";
import { Question } from "./Question";
import { QuestionsParser } from "./QuestionsParser";

export class Controller {
    generateService: GenerateDocumentService;
    moodleExporter: MoodleExporter;

    constructor(){
        this.generateService = new GenerateDocumentService();
        this.moodleExporter = new MoodleExporter();
    }

    generateDocument(
        documentName: string,
        variantsCount: number,
        questionsPerVariantCount: number,
        pagesPerTestCount: number,
        stats
    ) {
        Logger.log('test');
        const docName = documentName ?? 'col-test';
        const varsCount = variantsCount ?? 50;
        const questionsCount = questionsPerVariantCount ?? 15;
        const pagesPerTest = pagesPerTestCount ?? 4;

        const document = DocumentApp.create(docName);

        try {
            this.generateService.generate(document, varsCount, questionsCount, pagesPerTest, stats);
            document.saveAndClose();

            const generatedId = document.getId();
            const generatedFile = DriveApp.getFileById(generatedId);
            const directParents = DriveApp.getFileById(
                DocumentApp.getActiveDocument().getId()
            ).getParents();

            while (directParents.hasNext()) {
                generatedFile.moveTo(directParents.next());
            }

            var recipient = Session.getActiveUser().getEmail();
            var subject = 'Колоквіум ' + docName + ' готовий';
            var body =
                'Лінк на тести ' +
                documentName +
                '\n' +
                generatedFile.getUrl() +
                '\n';
            MailApp.sendEmail(recipient, subject, body);

            return generatedFile.getUrl();
        } catch (e) {
            Logger.log('Exception occured ', e);
            DocumentApp.getUi().alert('Помилка при генерації документу: ' + e);
            DriveApp.getFileById(document.getId()).setTrashed(true);
        }
    }

    exportMoodle(){
        var questions: Question[];

        try {
          questions = QuestionsParser.parse()
        }  catch(e) {
          Logger.log("Exception occured ", e)
          DocumentApp.getUi().alert("Помилка при генерації документу: " + e)
          return ;
        }

        const doc = this.moodleExporter.export(questions);

        const xml = XmlService.getPrettyFormat().format(doc)

        const curDoc = DocumentApp.getActiveDocument()
        const docFile = DriveApp.getFileById(curDoc.getId())
        const name = docFile.getName() + ".xml"
        const firstDir = docFile.getParents().next()

        if (!firstDir) {
          DocumentApp.getUi().alert('Не вдалось визначити батьківську директорію файлу')
          return ;
        }

        const generatedFile = firstDir.createFile(name, xml)
        const url = generatedFile.getUrl()

        DocumentApp.getUi()
          .alert('Згенерований файл: ' + url)
    }

    getStats(){
        let questions: Question[];
        try {
            questions = QuestionsParser.parse(); //parseQuestions()
        } catch (e) {
            Logger.log('Exception occured ', e);
            DocumentApp.getUi().alert('Помилка при генерації документу: ' + e);
            return { total: 0, topics: [] };
        }

        var topics = {};

        questions.forEach((q) =>
            q.topics.forEach((topic) => {
                topics[topic] = (topics[topic] ?? 0) + 1;
            })
        );

        var list: { topic: string; count: number }[] = [];

        for (const [topic, count] of Object.entries(topics)) {
            list.push({ topic, count: <number>count });
        }

        list.sort((a, b) => b.count - a.count);

        return {
            total: questions.length,
            topics: list,
        };
    }
}
