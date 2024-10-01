export class Question {

    public id: number;
    public text: string;
    public options: string[] = [];
    public usedTimes = 0;
    public incompatible: number[] = [];
    public tags: number[] = [];
    public topics: string[] = [];
    public correctOptions: (number | string)[] = [1]; // reserved
    public index = 0;

    constructor(id: number, text: string) {
        this.id = id;

        this.text = replaceFunc(text, this).trim();
    }

    addOption(text: string) {
        this.options.push(replaceFunc(text, this));
    }

    printAndGetAnswer(no: number, body: GoogleAppsScript.Document.Body) {
        var text = this.text;

        if (text.indexOf('```') == -1) {
            body
                .appendParagraph(no + '. ' + this.text)
                .setSpacingBefore(12)
                .asText()
                .setBold(true)
                .setFontSize(12)
        } else {
            // We have code formatted as ```code```
            const paragraph = body.appendParagraph(no + '. ');
            paragraph.setSpacingBefore(12).asText().setBold(true).setFontSize(12);
            var pos = 0;

            while (pos < text.length) {
                var start = text.indexOf('```', pos);

                if (start < 0) {
                    paragraph
                        .appendText(text.substring(pos))
                        .setBold(true)
                        .setFontSize(12);
                    pos = text.length;
                } else {
                    var startText = text.substring(pos, start);

                    if (startText) {
                        paragraph
                            .appendText(startText)
                            .setBold(true)
                            .setFontSize(12);
                    }

                    var end = text.indexOf('```', start + 3); // 3 is len of ```
                    var endText = text.substring(start + 3, end);

                    if (endText) {
                        paragraph
                            .appendText(endText)
                            .setBold(false)
                            .setFontSize(11)
                            .setFontFamily('Consolas');
                    } else {
                        pos = text.length;
                    }

                    pos = end + 3;
                }
            }
        }

        return this.printOptions(body);
    }

    textAsHtml() {
        var text = this.text;
        if (text.indexOf('```') == -1) return `<p>${text}</p>`;

        // We have code formatted as ```code```
        var out = '';
        var pos = 0;

        while (pos < text.length) {
            var start = text.indexOf('```', pos);

            if (start < 0) {
                out += `<p>${text.substring(pos)}</p>`;
                pos = text.length;
            } else {
                var startText = text.substring(pos, start);

                if (startText) {
                    out += `<p>${startText}</p>`;
                }

                var end = text.indexOf('```', start + 3); // 3 is len of ```
                var endText = text.substring(start + 3, end);

                if (endText) {
                    out += `<code>${endText}</code>`;
                } else {
                    pos = text.length;
                }

                pos = end + 3;
            }
        }

        return (<any>out)
            .replaceAll('\n', '<br/>')
            .replaceAll('<br/></p>', '</p>');
    }

    regenerateIndex() {
        this.index = Math.ceil(Math.random() * 100) + 35 * this.usedTimes;
    }

    printOptions(body: GoogleAppsScript.Document.Body): string[] {
        if (this.options.length == 0)
            throw 'Питання не має жодного варіанту відповіді: ' + this.text;

        var opts = [...this.options];
        var multipleOpts = opts.length > 1;
        var multipleCorrectOptions = this.correctOptions.length > 1;
        var result: string[] = ['-'];

        if (multipleOpts) {
            const correctOpts = this.correctOptions.map(
                (i) => opts[<number>i - 1]
            );
            // randomize
            opts.sort(function () {
                return Math.random() - 0.5;
            });

            // find correct option indexes
            result = correctOpts.map((co) => (opts.indexOf(co) + 1) + '');
        }

        for (let i = 0; i < opts.length; i++){
            this.printOpt(
                opts[i].trim(),
                i,
                body,
                multipleOpts,
                multipleCorrectOptions
            );
        }

        return result;
    }

    printOpt(o: string, idx: number, body: GoogleAppsScript.Document.Body, multipleOpts: boolean, multipleCorrectOptions: boolean) {
        // 25EF - ◯, 2610 - ☐
        var checkboxSymb = multipleCorrectOptions
            ? String.fromCharCode(parseInt('2610', 16))
            : String.fromCharCode(parseInt('25EF', 16));
        var text = multipleOpts ? checkboxSymb + ' ' + (idx + 1) + ') ' + o : o;

        if (multipleOpts) {
            body.appendParagraph(text).asText().setBold(false).setFontSize(11);
        } else {
            body.appendParagraph(text).asText().setBold(false).setFontSize(11);
        }
    }
}

function replaceFunc(text: string, question: Question): string {
    return text.replace(/@(\w+)\((.*)\)/g, function (match, fname, argsStr) {
        try {
            Logger.log(
                'Q' +
                    question.id +
                    ' : Replacing "' +
                    match +
                    ', fname=' +
                    fname +
                    ', args=' +
                    argsStr
            );

            var args = argsStr.split(',');

            switch (fname) {
                case 'TextField':
                    var lines = args[0] ? parseInt(args[0]) : 1;
                    var lineTpl =
                        '___________________________________________________________________________';
                    var text = lineTpl;
                    for (var i = 1; i < lines; i++) {
                        text += '\n' + lineTpl;
                    }
                    return text;
                case 'NotWith':
                    args.forEach(function (e) {
                        question.incompatible[parseInt(e.trim())] = 1;
                    });
                    return '';
                case 'Tag':
                    args.forEach(function (e) {
                        question.tags.push(parseInt(e.trim()));
                    });
                    return '';
                case 'MultipleOptions':
                    question.correctOptions = args.map((o) =>
                        parseInt(o.trim())
                    );
                    return '';
                case 'Topic':
                    args.forEach(function (a) {
                        question.topics.push(a.trim());
                    });
                    return '';
                default:
                    Logger.log(
                        'Q' +
                            question.id +
                            ' : Replacing "' +
                            match +
                            ' – Not matched a function'
                    );
                    return match;
            }
        } catch (err) {
            Logger.log(err);
            return match;
        }
    });
}
