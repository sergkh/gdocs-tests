import { QuestionsParser } from './QuestionsParser';
import GoogleDocument = GoogleAppsScript.Document.Document;
import { Question } from './Question';

export class GenerateDocumentService {
    generate(
        document: GoogleDocument,
        variantsCount,
        questionsCount,
        pagesPerTest,
        stats
    ) {
        var questions: Question[] = QuestionsParser.parse();
        
        // Form document
        if (questionsCount > questions.length / 2)
            throw (
                'Кількість запитань (' +
                questions.length +
                ') недостатня для формування тестів. Необхідно хоча б ' +
                questionsCount * 2 +
                ' запитань'
            );

        var answers: string[][][] = [];
        var outBody = document.getBody();

        for (var varNo = 1; varNo <= variantsCount; varNo++) {
            answers.push(
                this.formVariant(
                    varNo,
                    outBody,
                    this.mixQuestions(varNo, questions, questionsCount)
                )
            );
        }

        // Report with answers

        outBody
            .appendParagraph('Відповіді:')
            .setHeading(DocumentApp.ParagraphHeading.HEADING3);

        answers.forEach(function (variant, varIdx) {
            outBody
                .appendParagraph('Варіант: ' + (varIdx + 1) + ':')
                .asText()
                .setBold(true);

            var answText = '';

            variant.forEach(function (a, idx) {
                answText += idx + 1 + ' : ' + a.join(', ') + '\n';
            });

            outBody
                .appendParagraph(answText)
                .asText()
                .setFontSize(10)
                .setBold(false);
        });

        if (stats) {
            var statPar = outBody.appendParagraph('Статистика:');
            statPar.setHeading(DocumentApp.ParagraphHeading.HEADING1);

            var statText = '';

            questions.forEach(function (e, idx) {
                statText +=
                    'Запитання: ' +
                    (idx + 1) +
                    ' використано ' +
                    e.usedTimes +
                    ' раз\n';
            });

            outBody
                .appendParagraph(statText)
                .asText()
                .setBold(false)
                .setFontSize(10);
        }
    }

    // Builds a single variant section
    formVariant(number, body, questions: Question[]) {
        var titlePar = body.appendParagraph('Варіант ' + number);
        titlePar.setHeading(DocumentApp.ParagraphHeading.HEADING3);
        titlePar.setBold(true);
        titlePar.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        titlePar.setFontFamily('Arial');

        var namePar = body.appendParagraph(
            'ПІБ ______________________________________________   Група ____________'
        );
        namePar.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        namePar.setBold(false);
        namePar.setFontSize(12);

        var answers: string[][] = [];

        for (let i = 0; i < questions.length; i++) {
            var rightAnswer = questions[i].printAndGetAnswer(i + 1, body);
            answers.push(rightAnswer);
        }

        body.appendPageBreak();

        return answers;
    }

    mixQuestions(varNo, questions: Question[], questionsCount) {
        var clone = questions.slice(0);

        clone.forEach(function (el) {
            el.regenerateIndex();
        });

        // mix questions simultaneously trying to put most used to the end.
        clone.sort(function (q1, q2) {
            return q1.index - q2.index;
        });

        var copy: Question[] = [];

        for (let itm of clone){
            if (copy.length == questionsCount) continue;

            if (!this.containsIncompatible(copy, itm)) {
                copy.push(itm);
                itm.usedTimes++;
            }
        }

        if (copy.length < questionsCount)
            throw 'Кількість запитань менша за необхідну, можливо занадто багато несумісних питань';

        return copy;
    }

    containsIncompatible(arr, itm) {
        for (var i = 0, len = arr.length; i < len; i++) {
            if (arr[i].incompatible[itm.id] == 1) return true;
            if (this.hasCommonElements(arr[i].tags, itm.tags)) return true;
        }
        return false;
    }

    hasCommonElements(array1, array2) {
        for (var i = 0, len = array1.length; i < len; i++) {
            if (array2.indexOf(array1[i]) != -1) return true;
        }

        return false;
    }
}
