/**
 * Класс для построения коммерческих предложений.
 */
class BusinessProposalBuilder {

  /**
   * Создает экземпляр BusinessProposalBuilder.
   * @throws {Error} Если листы не найдены.
   */
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.ordersSheet = this.ss.getSheetByName(ORDERS_SHEET);
    this.productsSelectionSheet = this.ss.getSheetByName(PRODUCTS_SELECTION_SHEET);
    this.previewSheet = this.ss.getSheetByName(PREVIEW_SHEET);
    this.businessProposalSheet = this.ss.getSheetByName(BUSINESS_PROPOSAL_SHEET);

    if (!this.ordersSheet || !this.productsSelectionSheet || !this.previewSheet || !this.businessProposalSheet) {
      throw new Error("Листы не найдены. Пожалуйста, убедитесь, что названия листов указаны верно.");
    }

    this.yuanExchangeRate = this.productsSelectionSheet.getRange('D1').getValue();
    this.totalCost = 0;
    this.fastTotalDeliveryCost = 0;
    this.slowTotalDeliveryCost = 0;

    this.currentClientName = this.productsSelectionSheet.getRange('B1').getValue();
  }

  /**
   * Обработчик события редактирования.
   * @param {Object} e -- Объект события редактирования.
   */
  onEdit(e) {
    const {range} = e;

    // Проверка, что редактирование произошло в ячейке B1.
    if (range.getA1Notation() !== 'B1') {
      return;
    }

    const client = range.getValue(); // Получение значения клиента из ячейки B1

    // Получение данных из листа заказов
    const data = this
      .ordersSheet.getRange(ORDERS_START_POSITION, 1, this.ordersSheet.getLastRow(), this.ordersSheet.getLastColumn())
      .getValues();


    // Проверка на наличие данных
    if (!data || data.length === 0) {
      throw new Error("Данные заказов не найдены.");
    }

    this.clearSheet(this.productsSelectionSheet, 2);

    // Поиск заказов, связанных с клиентом
    const clientIndexes = [];

    data.forEach((row, i) => {
      const trimmedRowValue = String(row[2]).trim();
      if (trimmedRowValue === client) {
        clientIndexes.push(ORDERS_START_POSITION + i);
      }
    });

    // Проверка на наличие заказов для выбранного клиента
    if (clientIndexes.length === 0) {
      throw new Error("Заказы для выбранного клиента не найдены.");
    }

    // Копирование заказов для выбранного клиента на лист выбора товаров
    clientIndexes.forEach(index => {
      const currentRow = this.productsSelectionSheet.getLastRow() + 1;

      this.ordersSheet
        .getRange(index, 1, 1, this.ordersSheet.getLastColumn())
        .copyTo(this.productsSelectionSheet.getRange(currentRow, 1));

      // Вставка чекбокса "yes" в новую строку и установка его состояния как "выбран"
      this.productsSelectionSheet
        .getRange(currentRow, 1)
        .insertCheckboxes('yes')
        .check();

    });
  }

  /**
   * Создает коммерческое предложение.
   *
   * @throws {Error} Если данные заказов не найдены.
   * @throws {Error} Если не выбрано ни одной галочки.
   */
  build() {

    let number = 1; // Счетчик заказов на листе КП

    // Все заказы клиента на странице 'Выбор товаров'
    const clientOrders = this.productsSelectionSheet
      .getDataRange()
      .getValues()
      .slice(1);

    if (!clientOrders || clientOrders.length === 0) {
      throw new Error("Данные заказов не найдены.");
    }

    const hasCheckboxesSelected = clientOrders.some(([checked]) => checked === 'yes');

    if (!hasCheckboxesSelected) {
      throw new Error("Необходимо поставить хотя бы одну галочку");
    }

    // Очистка старых данных на листе КП и листе превью
    this.clearSheet(this.businessProposalSheet, BUSINESS_PROPOSAL_HEADER_ROW);
    this.clearSheet(this.previewSheet, 34);

    // Устанавливает сплошные белые границы вокруг ячеек  на листе "Готовое КП"
    // this.businessProposalSheet.getRange('A:G')
    //   .setBorder(true, true, true, true, true, true, 'white', SpreadsheetApp.BorderStyle.SOLID);

    // Копирование шапки на лист КП
    this.previewSheet.getRange(HEADER_POSITION)
      .copyTo(this.businessProposalSheet.getRange(BUSINESS_PROPOSAL_HEADER_ROW, 1));

    // Копирование выбранных заказов
    clientOrders.forEach((order, index) => {

      // Если галочки нет — переходим к следующему товару
      if (order[0] !== 'yes') {
        return;
      }

      // Копирование заказа на лист "Лист превью" ???
      this.productsSelectionSheet.getRange(index + 2, 2, 1, this.productsSelectionSheet.getLastColumn())
        .copyTo(this.previewSheet.getRange(this.previewSheet.getLastRow() + 1, 2));


      // Формирование объекта с данными заказа
      const orderData = {
        number: number++,
        amount: order[letterToIndex('N')],
        volume: order[letterToIndex('S')],
        weight: order[letterToIndex('T')],
        sheathingWeight: order[letterToIndex('U')],
        totalWeight: order[letterToIndex('V')],
        cargoRate14_21: order[letterToIndex('AA')],
        cargoRate35_45: order[letterToIndex('AB')],
        costOfCargoPackaging: order[letterToIndex('AD')],
        unloadingCost: order[letterToIndex('AE')],
        insurance: order[letterToIndex('AT')],
        unitCostIncludingCommission: order[letterToIndex('R')],
        priceOfDeliveryInChinaPerBatch: order[letterToIndex('O')],
        fastFreightCost: order[letterToIndex('AL')],
        slowFreightCost: order[letterToIndex('AM')],
        purchase: order[letterToIndex('AK')],
      };

      // Заполнение шаблонного заказа данными
      const entries = Object.entries(orderData);
      for (const [position, data] of entries) {
        this.previewSheet.getRange(productTemplateCell[position]).setValue(data);
      }

      // Копирование изображения
      this.productsSelectionSheet.getRange(`D${index + 2}`)
        .copyTo(this.previewSheet.getRange(productTemplateCell.imageRange).merge());

      let businessProposalSheetStartRow = this.businessProposalSheet.getLastRow() + 1;

      // Копирование заказа на лист КП
      this.previewSheet.getRange(TEMPLATE_PRODUCT_RANGE)
        .copyTo(this.businessProposalSheet.getRange(businessProposalSheetStartRow + 1, 1))

      // Собирает переносы страниц
      if ((orderData.number) % 2 === 0) {
        pagebreaks.push(this.businessProposalSheet.getLastRow());
      }

      // Устанавливает красную пунктирную границу
      this.businessProposalSheet
        .getRange(this.businessProposalSheet.getLastRow(), this.businessProposalSheet.getLastColumn())
        .setBorder(true, null, true, null, null, null, 'red', SpreadsheetApp.BorderStyle.DOTTED);

      // Аккумулирование итоговых данных по КП
      const {purchase, fastFreightCost, slowFreightCost} = orderData;
      this.totalCost += purchase;
      this.fastTotalDeliveryCost += fastFreightCost;
      this.slowTotalDeliveryCost += slowFreightCost;

      // Очистка шаблона заказа
      for (const [position, data] of entries) {
        this.previewSheet.getRange(productTemplateCell[position]).clearContent();
      }
      this.previewSheet.getRange(productTemplateCell.imageRange).clearContent(); // Очистка изображения
    });

    // Копирование итоговых данных в шапку на листе КП
    const headerData = [
      this.yuanExchangeRate,
      this.totalCost + this.totalCost * 0.05,
      this.fastTotalDeliveryCost,
      this.slowTotalDeliveryCost,
    ]

    this.businessProposalSheet
      .getRangeList(Object.values(headerTemplateCell))
      .getRanges()
      .forEach((range, i) => range.setValue(headerData[i]));

    // Копирование подвала на лист КП
    const businessProposalSheetLastRow = this.businessProposalSheet.getLastRow() + 2;
    this.previewSheet.getRange(FOOTER_RANGE)
      .copyTo(this.businessProposalSheet.getRange(businessProposalSheetLastRow, 1));

    // Обновляет перенос последний страницы
    if ((number - 1) % 2 === 0) {
      pagebreaks[pagebreaks.length - 1] = this.businessProposalSheet.getLastRow()
    } else {
      pagebreaks.push(this.businessProposalSheet.getLastRow());
    }

    userProperties.setProperties({
      "pagebreaks": pagebreaks.join(','),
    });

  }

  /**
   * Очистка листа, начиная с определенной строки.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист для очистки.
   * @param {number} fromRow - Номер строки, с которой начать очистку.
   */
  clearSheet(sheet, fromRow = 1) {
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      return;
    }

    const oldContentRange = sheet
      .getRange(fromRow, 1, lastRow, sheet.getLastColumn());

    oldContentRange
      .clearFormat()
      .clearDataValidations()
      .clear();
  }

  /**
   * Генерирует имя папки на основе текущей даты и имени клиента.
   *
   * @returns {string} Сгенерированное имя папки.
   */
  createFolderName() {
    const [date] = new Date().toISOString().replace(/T/, ' ').replace(/:/g, '-').split('.');
    return `${date}_${this.currentClientName}`
  }

  openModal() {
    const service = HtmlService.createTemplateFromFile("src/frontend/dialog");
    const htmlOutput = service.evaluate()

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Сохранить КП');
  }

  /**
   * Возвращает массив URL-адресов для экспорта разделов документа в формате PDF.
   *
   * @returns {string[]} Массив URL-адресов для экспорта разделов в формате PDF.
   */
  getPdfUrls() {
    const pagebreaks = userProperties.getProperty("pagebreaks").split(',').map(pb => parseInt(pb, 10));
    const urls = [];

    const firstColumn = 0;
    const lastColumn = this.businessProposalSheet.getLastColumn();

    for (let i = 1; i < pagebreaks.length; ++i) {
      const firstRow = pagebreaks[i - 1] + 1;
      const lastRow = pagebreaks[i];

      const url = "https://docs.google.com/spreadsheets/d/" + this.ss.getId() + "/export" +
        "?format=pdf&" +
        "size=A4&" +
        "fzr=true&" +
        "portrait=true&" +
        "fitw=true&" +
        "gridlines=false&" +
        "printtitle=false&" +
        "top_margin=0.5&" +
        "bottom_margin=0.25&" +
        "left_margin=0.5&" +
        "right_margin=0.5&" +
        "sheetnames=false&" +
        "pagenum=UNDEFINED&" +
        "attachment=true&" +
        "gid=" + this.businessProposalSheet.getSheetId() + '&' +
        "r1=" + firstRow + "&c1=" + firstColumn + "&r2=" + lastRow + "&c2=" + lastColumn;
      urls.push(url);
    }
    return urls;
  }

  /**
   * Возвращает URL-адрес для экспорта листа КП в формате XLSX.
   *
   * @returns {string} URL-адрес для экспорта в формате XLSX.
   */
  getXlsxUrl() {
    return "https://docs.google.com/spreadsheets/d/" + this.ss.getId() + "/export" + "?gid=" + this.businessProposalSheet.getSheetId();
  }

  /**
   * Генерирует файлы и сохраняет их в указанную папку.
   * @param {string} folderName - Название папки для сохранения файлов.
   * @returns {Object} - Объект с информацией о созданных файлах и папке.
   */
  async generateFiles(folderName) {
    const token = ScriptApp.getOAuthToken();
    const fileName = folderName.split('_').reverse().join('_');
    const folder = getFolderByName_(folderName);
    const pdfUrls = this.getPdfUrls();
    const xlsxUrl = this.getXlsxUrl();
    const params = {method: "GET", headers: {"authorization": "Bearer " + token}};

    const urls = [xlsxUrl, ...pdfUrls];

    const requests = urls.map(url => {
      return {url, ...params}
    });

    const responses = UrlFetchApp.fetchAll(requests);
    const blobs = responses.map(response => response.getBlob());

    const xlsxBlob = blobs[0].setName(`${fileName}.xlsx`);
    const pdfBlobs = blobs.slice(1);

    const pdfBlob = await mergeAllPDFs(pdfBlobs, `${fileName}.pdf`);
    const xlsx = folder.createFile(xlsxBlob)
    const pdf = folder.createFile(pdfBlob);

    return {
      pdfId: pdf.getId(),
      xlsxId: xlsx.getId(),
      fileName,
      folder,
    };
  }

}
