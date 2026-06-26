const result = $input.first().json;
if (!result.success) {
  throw new Error('Ошибка обработки Excel: ' + JSON.stringify(result));
}

// Метаданные из ноды "Подготовить данные"
const meta = $('Подготовить данные').first().json;

const outBase64 = result.fileBase64;
const processedFileName = `обработан_${meta.attachFileName}`;

const summary =
  `Обработано строк: ${result.totalProcessed}\n` +
  `🔴 Запрещённые: ${result.prohibited.length}\n` +
  `🟡 Превысили $200: ${result.violators.length}`;

// Сохранить состояние в static data (без fs)
const state = $getWorkflowStaticData('global');
state.registryState = {
  senderEmail: meta.senderEmail,
  emailSubject: meta.emailSubject,
  violators: result.violators,
  prohibited: result.prohibited,
  processedFileName,
  fileBase64: outBase64,
  processedAt: new Date().toISOString(),
  usdRate: meta.usdRate,
  status: 'pending_review',
};

return [{
  json: {
    senderEmail: meta.senderEmail,
    emailSubject: meta.emailSubject,
    violators: result.violators,
    prohibited: result.prohibited,
    processedFileName,
    summary,
    usdRate: meta.usdRate,
    processedAt: state.registryState.processedAt,
    status: 'pending_review',
  },
  binary: {
    processedFile: {
      data: outBase64,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileName: processedFileName,
    },
  },
}];
