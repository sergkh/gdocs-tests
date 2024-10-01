import { Controller } from './Controller';

/**
 * The Script that generates test variants from the list of test questions stored in the Google Document.
 *
 * Extending Google Docs developer guide:
 *     https://developers.google.com/apps-script/guides/docs
 *
 * Document service reference documentation:
 *     https://developers.google.com/apps-script/reference/document/
 */
export function onOpen() {
    DocumentApp.getUi()
        .createMenu('Колоквіум')
        .addItem('Генератор...', 'generateDialog')
        .addItem('Статистика...', 'topicsDialog')
        .addItem('Екпорт в Moodle', 'moodleExport')
        .addToUi();
}

export function generateDialog() {
    var html = HtmlService.createTemplateFromFile('Form').evaluate();

    DocumentApp.getUi().showModalDialog(html, 'Генерація колоквіуму');
}

export function topicsDialog() {
    var html = HtmlService.createTemplateFromFile('Topics').evaluate();

    DocumentApp.getUi().showModalDialog(html, 'Звіт по темам');
}

export function moodleExport() {
    const controller = new Controller();
    controller.exportMoodle();
}

export function generateDocument(
    documentName,
    variantsCount,
    questionsPerVariantCount,
    pagesPerTestCount,
    stats
) {
    const controller = new Controller();
    return controller.generateDocument(
        documentName,
        variantsCount,
        questionsPerVariantCount,
        pagesPerTestCount,
        stats
    );
}

export function questionsStats() {
    const controller = new Controller();
    return controller.getStats();
}
