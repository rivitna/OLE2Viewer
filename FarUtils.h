//---------------------------------------------------------------------------
#ifndef __FARUTILS_H__
#define __FARUTILS_H__
//---------------------------------------------------------------------------
namespace Far {
//---------------------------------------------------------------------------
// Инициализация информации Far
void Init(
  const struct PluginStartupInfo *pInfo
  );
// Создание объекта "настройки" плагина
HANDLE CreateSettingsObj();
// Удаление объекта "настройки" плагина
void FreeSettingsObj(
  HANDLE hSettings
  );
// Загрузка значения настройки плагина
unsigned __int64 LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  unsigned __int64 nDefault
  );
// Загрузка значения настройки плагина
const wchar_t* LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const wchar_t *pwszDefault
  );
// Загрузка значения настройки плагина
size_t LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  void *pValue,
  size_t nSize
  );
// Сохранение значения настройки плагина
bool SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  unsigned __int64 nValue
  );
// Сохранение значения настройки плагина
bool SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const wchar_t *pwszValue
  );
// Сохранение значения настройки плагина
bool SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const void *pValue,
  size_t nSize
  );
// Получение строки сообщения из языкового файла
const wchar_t* GetLocMsg(
  int msgId
  );
// Отображение сообщения
void ShowMessage(
  int nFlags,
  const wchar_t *pwszCaption,
  const wchar_t *pwszText,
  const wchar_t *pwszErrorItem
  );
// Отображение сообщения
void ShowMessage(
  int nFlags,
  int captionId,
  int textId,
  const wchar_t *pwszErrorItem
  );
// Отображение меню
intptr_t DisplayMenu(
  const wchar_t *pwszTitle,
  const struct FarMenuItem *pItems,
  size_t nNumItems
  );
// Отображение диалогового окна конфигурации плагина
bool ShowConfigDialog(
  int *pnEnablePlugin,
  int *pnUseExtensionFilter
  );
// Отображение диалогового окна подтверждения извлечения файлов
bool ShowConfirmExtractDialog(
  wchar_t *pwszDestPathBuf,
  size_t nDestPathBufSize
  );
// Обновление содержимое панели
void UpdatePanel();
// Перерисовка панели
void RedrawPanel(
  size_t nCurrentItem,
  size_t nTopPanelItem
  );
// Сохранение экрана
HANDLE SaveScreen();
// Восстановление области экрана
void RestoreScreen(
  HANDLE hScreen
  );
// Установка типа и состояния индикатора прогресса
void SetProgressState(
  int nState
  );
// Получение общей информации об активной панели
bool GetPanelInfo(
  struct PanelInfo *pPanelInfo
  );
// Получение текущего каталога активной панели
size_t GetPanelDirectory(
  struct FarPanelDirectory *pPanelDir,
  size_t nBufSize
  );
// Получение текущего файлового элемента активной панели
size_t GetCurrentPanelItem(
  struct FarGetPluginPanelItem *pItem
  );
// Получение пути к файлу для текущего файлового элемента активной панели
// (Возвращенный указатель необходимо освободить при помощи функции free)
wchar_t* GetCurrentPanelItemFilePath();
// Получение полного пути
// (Возвращенный указатель необходимо освободить при помощи функции free)
wchar_t* ExpandPath(
  const wchar_t *pwszPath
  );
// Удаление из строки начальных и конечных пробелов
wchar_t* TrimStr(
  wchar_t *s
  );
// Удаление из строки всех двойных кавычек
wchar_t* UnquoteStr(
  wchar_t *s
  );
//---------------------------------------------------------------------------
}  // namespace Far
//---------------------------------------------------------------------------
#endif  // __FARUTILS_H__
