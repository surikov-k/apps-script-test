const userProperties = PropertiesService.getUserProperties();

const ORDERS_SHEET = 'Заказы'; // Название листа с заказами
const PRODUCTS_SELECTION_SHEET = 'Выбор товаров'; // Название листа с выбором товаров
const PREVIEW_SHEET = 'Лист превью'; // Название листа с превью
const BUSINESS_PROPOSAL_SHEET = 'Готовое КП'; // Название листа с КП

const ORDERS_START_POSITION = 200; // Позиция, с которой начинаются заказы на листе 'Заказы'

// Позиции шаблонов продукта, шапки и подвала на листе превью
const HEADER_POSITION = 'B1:H12'; // Диапазон шаблона шапки
const TEMPLATE_PRODUCT_RANGE = 'B14:H30' // Диапазон шаблона продукта
const FOOTER_RANGE = 'C31:I31' // Диапазон подвала

const BUSINESS_PROPOSAL_HEADER_ROW = 3 // Ряд, на котором начинается шапка на листе КП

const productTemplateCell = {
  number: 'B14',
  imageRange: 'C15:C28',
  amount: 'G14',
  volume: 'H16',
  weight: 'H17',
  sheathingWeight: 'H18',
  totalWeight: 'H19',
  cargoRate14_21: 'H21',
  cargoRate35_45: 'H22',
  costOfCargoPackaging: 'H23',
  unloadingCost: 'H24',
  insurance: 'H25',
  unitCostIncludingCommission: 'H26',
  priceOfDeliveryInChinaPerBatch: 'H27',
  fastFreightCost: 'H28',
  slowFreightCost: 'H29',
  purchase: 'H30'
}

const headerTemplateCell = {
  yuanExchangeRate: 'C3',
  // totalCost: 'C5',
  totalCostWith5Percent: 'C6',
  fastTotalDeliveryCost: 'C8',
  slowTotalDeliveryCost: 'C9',
}

pagebreaks = [1];

const businessProposalBuilder = new BusinessProposalBuilder();

/**
 * Обработчик события редактирования.
 * @param {Object} e - Объект события редактирования.
 */
function onEdit(e) {
  try {
    businessProposalBuilder.onEdit(e);
  } catch (error) {
    Browser.msgBox(
      'Внимание',
      `Произошла ошибка: ${error.message}`,
      Browser.Buttons.OK
    );
  }
}

/**
 * Обработчик клика на кнопку 'Сформировать КП'.
 */
function buildBusinessProposal() {
  try {
    businessProposalBuilder.build()
    businessProposalBuilder.ss.setActiveSheet(businessProposalBuilder.businessProposalSheet);
  } catch (error) {
    Browser.msgBox(
      'Внимание',
      `Произошла ошибка: ${error.message}`,
      Browser.Buttons.OK
    );
  }
}

/**
 * Обработчик клика на кнопку 'Сохранить в PDF'.
 */
function openModal() {
  businessProposalBuilder.openModal();
}

/**
 * Сохраняет файлы на устройство.
 *
 * @async
 * @function
 * @returns {string} Ссылка на архивную папку с сохраненными файлами.
 */
async function saveToDevice() {
  const folderName = businessProposalBuilder.createFolderName();
  const {pdfId, xlsxId, fileName} = await businessProposalBuilder.generateFiles(folderName);

  return {
    pdfId,
    xlsxId,
    fileName,
    apiKey: ScriptApp.getOAuthToken()
  };
}

/**
 * Сохраняет файлы в Google Drive.
 *
 * @async
 * @function
 * @returns {Object} Объект с именем и ссылкой на сохраненную папку в Google Drive.
 */
async function saveToDrive() {
  const folderName = businessProposalBuilder.createFolderName();
  const {folder} = await businessProposalBuilder.generateFiles(folderName);

  return {
    name: folder.getName(),
    link: folder.getUrl()
  };
}

function develop() {
}



