//---------------------------------------------------------------------------
#ifndef __PLUGLANG_H__
#define __PLUGLANG_H__
//---------------------------------------------------------------------------
// Идентификаторы сообщений
enum
{
  MSG_PLUGIN_NAME,
  // Конфигурация
  MSG_CONFIG_ENABLE,
  MSG_CONFIG_USEEXTFILTER,
  // Информационная панель
  MSG_INFOL_VERSION,
  MSG_INFOL_TYPE,
  MSG_INFOL_SIZE,
  MSG_INFOL_STREAMS,
  MSG_INFOL_STORAGES,
  MSG_INFOL_MACROS,
  // Открытие
  MSG_OPEN_FILE_OPENING,
  MSG_OPEN_MENU_OPEN,
  // Извлечение
  MSG_EXTRACT_TITLE,
  MSG_EXTRACT_PATH_TEXT,
  // Кнопки
  MSG_BTN_OK,
  MSG_BTN_CANCEL,
  MSG_BTN_EXTRACT,
  // Ошибки
  MSG_ERROR_PARAM,
  MSG_ERROR_DATA,
  MSG_ERROR_MEM,
  MSG_ERROR_OPEN,
  MSG_ERROR_READ,
  MSG_ERROR_WRITE,
  MSG_ERROR_UNKNOWN
};
//---------------------------------------------------------------------------
#endif  // __PLUGLANG_H__
