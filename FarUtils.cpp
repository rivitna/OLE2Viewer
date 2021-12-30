//---------------------------------------------------------------------------
#include <stddef.h>
#include <malloc.h>
#include <memory.h>
#include "Far\plugin.hpp"
#include "Far\DlgBuilder.hpp"
#include "StrUtils.h"
#include "FileUtil.h"
#include "Guids.h"
#include "PlugLang.h"
#include "FarUtils.h"
//---------------------------------------------------------------------------
struct PluginStartupInfo g_FarSInfo;
struct FarStandardFunctions g_FSF;
//---------------------------------------------------------------------------
/***************************************************************************/
/* Init - Инициализация информации Far                                     */
/***************************************************************************/
void Far::Init(
  const struct PluginStartupInfo *pInfo
  )
{
  if (pInfo->StructSize < sizeof(struct PluginStartupInfo))
    return;
  memcpy(&g_FarSInfo, pInfo, sizeof(struct PluginStartupInfo));
  memcpy(&g_FSF, pInfo->FSF, sizeof(struct FarStandardFunctions));
  g_FarSInfo.FSF = &g_FSF;
}
/***************************************************************************/
/* CreateSettingsObj - Создание объекта "настройки" плагина                */
/***************************************************************************/
HANDLE Far::CreateSettingsObj()
{
  // Создание объекта "настройки"
  FarSettingsCreate settingsCreate;
  settingsCreate.StructSize = sizeof(settingsCreate);
  settingsCreate.Guid = OLE2VIEWER_GUID;
  settingsCreate.Handle = INVALID_HANDLE_VALUE;
  if (g_FarSInfo.SettingsControl(INVALID_HANDLE_VALUE, SCTL_CREATE,
                                 PSL_ROAMING, &settingsCreate))
    return settingsCreate.Handle;
  return INVALID_HANDLE_VALUE;
}
/***************************************************************************/
/* FreeSettingsObj - Удаление объекта "настройки" плагина                */
/***************************************************************************/
void Far::FreeSettingsObj(
  HANDLE hSettings
  )
{
  // Удаление объекта "настройки"
  g_FarSInfo.SettingsControl(hSettings, SCTL_FREE, 0, 0);
}
/***************************************************************************/
/* LoadSetting - Загрузка значения настройки плагина                       */
/***************************************************************************/
unsigned __int64 Far::LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  unsigned __int64 nDefault
  )
{
  // Загрузка элемента настроек плагина
  FarSettingsItem item;
  item.StructSize = sizeof(FarSettingsItem);
  item.Root = Root;
  item.Name = pwszName;
  item.Type = FST_QWORD;
  if (!g_FarSInfo.SettingsControl(hSettings, SCTL_GET, 0, &item))
    return nDefault;
  return item.Number;
}
/***************************************************************************/
/* LoadSetting - Загрузка значения настройки плагина                       */
/***************************************************************************/
const wchar_t* Far::LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const wchar_t *pwszDefault
  )
{
  // Загрузка элемента настроек плагина
  FarSettingsItem item;
  item.StructSize = sizeof(FarSettingsItem);
  item.Root = Root;
  item.Name = pwszName;
  item.Type = FST_STRING;
  if (!g_FarSInfo.SettingsControl(hSettings, SCTL_GET, 0, &item))
    return pwszDefault;
  return item.String;
}
/***************************************************************************/
/* LoadSetting - Загрузка значения настройки плагина                       */
/***************************************************************************/
size_t Far::LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  void *pValue,
  size_t nSize
  )
{
  // Загрузка элемента настроек плагина
  FarSettingsItem item;
  item.StructSize = sizeof(FarSettingsItem);
  item.Root = Root;
  item.Name = pwszName;
  item.Type = FST_DATA;
  if (!g_FarSInfo.SettingsControl(hSettings, SCTL_GET, 0, &item))
    return 0;
  if (pValue && (nSize >= item.Data.Size))
    memcpy(pValue, item.Data.Data, item.Data.Size);
  return item.Data.Size;
}
/***************************************************************************/
/* SaveSetting - Сохранение значения настройки плагина                     */
/***************************************************************************/
bool Far::SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  unsigned __int64 nValue
  )
{
  // Сохранение элемента настроек плагина
  FarSettingsItem item;
  item.StructSize = sizeof(FarSettingsItem);
  item.Root = Root;
  item.Name = pwszName;
  item.Type = FST_QWORD;
  item.Number = nValue;
  return (0 != g_FarSInfo.SettingsControl(hSettings, SCTL_SET, 0, &item));
}
/***************************************************************************/
/* SaveSetting - Сохранение значения настройки плагина                     */
/***************************************************************************/
bool Far::SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const wchar_t *pwszValue
  )
{
  // Сохранение элемента настроек плагина
  FarSettingsItem item;
  item.StructSize = sizeof(FarSettingsItem);
  item.Root = Root;
  item.Name = pwszName;
  item.Type = FST_STRING;
  item.String = pwszValue;
  return (0 != g_FarSInfo.SettingsControl(hSettings, SCTL_SET, 0, &item));
}
/***************************************************************************/
/* SaveSetting - Сохранение значения настройки плагина                     */
/***************************************************************************/
bool Far::SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const void *pValue,
  size_t nSize
  )
{
  // Сохранение элемента настроек плагина
  FarSettingsItem item;
  item.StructSize = sizeof(FarSettingsItem);
  item.Root = Root;
  item.Name = pwszName;
  item.Type = FST_DATA;
  item.Data.Size = nSize;
  item.Data.Data = pValue;
  return (0 != g_FarSInfo.SettingsControl(hSettings, SCTL_SET, 0, &item));
}
/***************************************************************************/
/* GetLocMsg - Получение строки сообщения из языкового файла               */
/***************************************************************************/
const wchar_t* Far::GetLocMsg(
  int msgId
  )
{
  return g_FarSInfo.GetMsg(&OLE2VIEWER_GUID, msgId);
}
/***************************************************************************/
/* ShowMessage - Отображение сообщения                                     */
/***************************************************************************/
void Far::ShowMessage(
  int nFlags,
  const wchar_t *pwszCaption,
  const wchar_t *pwszText,
  const wchar_t *pwszErrorItem
  )
{
  const wchar_t *msgItems[3];
  msgItems[0] = pwszCaption;
  msgItems[1] = pwszText;
  size_t nNumItems = 2;
  if (pwszErrorItem)
  {
    msgItems[2] = pwszErrorItem;
    nNumItems++;
  }
  g_FarSInfo.Message(&OLE2VIEWER_GUID, &GUID_O2V_MESSAGE_BOX, nFlags,
                     NULL, msgItems, nNumItems, 0);
}
/***************************************************************************/
/* ShowMessage - Отображение сообщения                                     */
/***************************************************************************/
void Far::ShowMessage(
  int nFlags,
  int captionId,
  int textId,
  const wchar_t *pwszErrorItem
  )
{
  // Отображение сообщения
  ShowMessage(nFlags, GetLocMsg(captionId), GetLocMsg(textId),
              pwszErrorItem);
}
/***************************************************************************/
/* DisplayMenu - Отображение меню                                          */
/***************************************************************************/
intptr_t Far::DisplayMenu(
  const wchar_t *pwszTitle,
  const struct FarMenuItem *pItems,
  size_t nNumItems
  )
{
  return g_FarSInfo.Menu(&OLE2VIEWER_GUID, &GUID_O2V_MENU, -1, -1, 0,
                         FMENU_WRAPMODE, pwszTitle, NULL, NULL, NULL, NULL,
                         pItems, nNumItems);
}
/***************************************************************************/
/* ShowConfigDialog - Отображение диалогового окна конфигурации плагина    */
/***************************************************************************/
bool Far::ShowConfigDialog(
  int *pnEnablePlugin,
  int *pnUseExtensionFilter
  )
{
  PluginDialogBuilder dlgBuilder(g_FarSInfo,
                                 OLE2VIEWER_GUID,
                                 GUID_O2V_CONFIG_DIALOG,
                                 MSG_PLUGIN_NAME,
                                 (const wchar_t *)0);

  dlgBuilder.AddCheckbox(MSG_CONFIG_ENABLE, pnEnablePlugin);
  dlgBuilder.AddCheckbox(MSG_CONFIG_USEEXTFILTER, pnUseExtensionFilter);
  dlgBuilder.AddOKCancel(MSG_BTN_OK, MSG_BTN_CANCEL, -1, true);

  if (dlgBuilder.ShowDialog())
    return true;
  return false;
}
/***************************************************************************/
/* ShowConfirmExtractDialog - Отображение диалогового окна подтверждения   */
/*                            извлечения файлов                            */
/***************************************************************************/
bool Far::ShowConfirmExtractDialog(
  wchar_t *pwszDestPathBuf,
  size_t nDestPathBufSize
  )
{
  PluginDialogBuilder dlgBuilder(g_FarSInfo,
                                 OLE2VIEWER_GUID,
                                 GUID_O2V_OTHER_DIALOG,
                                 MSG_EXTRACT_TITLE,
                                 (const wchar_t *)0);

  dlgBuilder.AddText(GetLocMsg(MSG_EXTRACT_PATH_TEXT));
  dlgBuilder.AddEditField(pwszDestPathBuf,
                          (int)nDestPathBufSize, 60)->Flags |= DIF_FOCUS;
  dlgBuilder.AddSeparator();
  dlgBuilder.AddOKCancel(MSG_BTN_EXTRACT, MSG_BTN_CANCEL, -1, true);

  if (dlgBuilder.ShowDialog())
    return true;
  return false;
}
/***************************************************************************/
/* UpdatePanel - Обновление содержимое панели                              */
/***************************************************************************/
void Far::UpdatePanel()
{
  // Обновление содержимое панели
  g_FarSInfo.PanelControl(PANEL_ACTIVE, FCTL_UPDATEPANEL, 0, NULL);
}
/***************************************************************************/
/* RedrawPanel - Перерисовка панели                                        */
/***************************************************************************/
void Far::RedrawPanel(
  size_t nCurrentItem,
  size_t nTopPanelItem
  )
{
  struct PanelRedrawInfo ri;
  ri.StructSize = sizeof(struct PanelRedrawInfo);
  ri.CurrentItem = nCurrentItem;
  ri.TopPanelItem = nTopPanelItem;
  // Перерисовка панели
  g_FarSInfo.PanelControl(PANEL_ACTIVE, FCTL_REDRAWPANEL, 0, &ri);
}
/***************************************************************************/
/* SaveScreen - Сохранение экрана                                          */
/***************************************************************************/
HANDLE Far::SaveScreen()
{
  // Сохранение области экрана
  return g_FarSInfo.SaveScreen(0, 0, -1, -1);
}
/***************************************************************************/
/* RestoreScreen - Восстановление области экрана                           */
/***************************************************************************/
void Far::RestoreScreen(
  HANDLE hScreen
  )
{
  // Восстановление области экрана
  g_FarSInfo.RestoreScreen(hScreen);
}
/***************************************************************************/
/* SetProgressState - Установка типа и состояния индикатора прогресса      */
/***************************************************************************/
void Far::SetProgressState(
  int nState
  )
{
  // Установка типа и состояния индикатора прогресса
  g_FarSInfo.AdvControl(&OLE2VIEWER_GUID, ACTL_SETPROGRESSSTATE,
                        (intptr_t)nState, NULL);
}
/***************************************************************************/
/* GetPanelInfo - Получение общей информации об активной панели            */
/***************************************************************************/
bool Far::GetPanelInfo(
  struct PanelInfo *pPanelInfo
  )
{
  pPanelInfo->StructSize = sizeof(struct PanelInfo);
  return (0 != g_FarSInfo.PanelControl(PANEL_ACTIVE, FCTL_GETPANELINFO, 0,
                                       pPanelInfo));
}
/***************************************************************************/
/* GetPanelDirectory - Получение текущего каталога активной панели         */
/***************************************************************************/
size_t Far::GetPanelDirectory(
  struct FarPanelDirectory *pPanelDir,
  size_t nBufSize
  )
{
  return (size_t)g_FarSInfo.PanelControl(PANEL_ACTIVE,
                                         FCTL_GETPANELDIRECTORY,
                                         nBufSize, pPanelDir);
}
/***************************************************************************/
/* GetCurrentPanelItem - Получение текущего файлового элемента активной    */
/*                       панели                                            */
/***************************************************************************/
size_t Far::GetCurrentPanelItem(
  struct FarGetPluginPanelItem *pItem
  )
{
  // Получение текущего файлового элемента панели
  return (size_t)g_FarSInfo.PanelControl(PANEL_ACTIVE,
                                         FCTL_GETCURRENTPANELITEM,
                                         0, pItem);
}
/***************************************************************************/
/* GetCurrentPanelItemFilePath - Получение пути к файлу для текущего       */
/*                               файлового элемента активной панели        */
/*                               (Возвращенный указатель необходимо        */
/*                                освободить при помощи функции free)      */
/***************************************************************************/
wchar_t* Far::GetCurrentPanelItemFilePath()
{
  struct PanelInfo panelInfo;

  // Получение общей информации об активной панели
  if (!GetPanelInfo(&panelInfo) ||
      (panelInfo.SelectedItemsNumber != 1) ||
      (panelInfo.PanelType != PTYPE_FILEPANEL))
    return NULL;

  size_t nBufSize;

  // Определение размера буфера для получения текущего файлового элемента
  // панели
  nBufSize = Far::GetCurrentPanelItem(NULL);
  if (nBufSize == 0)
    return NULL;

  struct FarGetPluginPanelItem item;
  item.StructSize = sizeof(struct FarGetPluginPanelItem);
  item.Size = nBufSize;
  struct PluginPanelItem *pPanelItem;
  pPanelItem = (struct PluginPanelItem *)malloc(nBufSize);
  if (!pPanelItem)
    return NULL;
  item.Item = pPanelItem;

  // Получение текущего файлового элемента панели
  Far::GetCurrentPanelItem(&item);

  wchar_t *pwszFilePath = NULL;

  if (!(pPanelItem->FileAttributes & FILE_ATTRIBUTE_DIRECTORY))
  {
    // Определение размера буфера для текущего каталога активной панели
    nBufSize = GetPanelDirectory(NULL, 0);
    if (nBufSize != 0)
    {
      struct FarPanelDirectory *pPanelDir;
      pPanelDir = (struct FarPanelDirectory *)malloc(nBufSize);
      if (pPanelDir)
      {
        // Получение текущего каталога активной панели
        pPanelDir->StructSize = sizeof(struct FarPanelDirectory);
        GetPanelDirectory(pPanelDir, nBufSize);

        if (pPanelDir->Name && pPanelDir->Name[0] &&
            IsRelativePath(pPanelItem->FileName))
        {
          // Объединение путей
          pwszFilePath = AllocAndCombinePath(pPanelDir->Name,
                                             pPanelItem->FileName);
        }
        else
        {
          // Копирование имени файла
          pwszFilePath = StrDup(pPanelItem->FileName);
        }

        free(pPanelDir);
      }
    }
  }

  free(pPanelItem);

  return pwszFilePath;
}
/***************************************************************************/
/* ExpandPath - Получение полного пути                                     */
/*              (Возвращенный указатель необходимо освободить при помощи   */
/*               функции free)                                             */
/***************************************************************************/
wchar_t* Far::ExpandPath(
  const wchar_t *pwszPath
  )
{
  // Определение размера буфера для полного пути
  size_t nFullPathSize = g_FSF.ConvertPath(CPM_FULL, pwszPath, NULL, 0);
  if (nFullPathSize == 0)
    return NULL;

  wchar_t *pwszFullPath = (wchar_t *)malloc(nFullPathSize * sizeof(wchar_t));
  if (!pwszFullPath)
    return NULL;

  // Получение полного пути
  g_FSF.ConvertPath(CPM_FULL, pwszPath, pwszFullPath, nFullPathSize);

  return pwszFullPath;
}
/***************************************************************************/
/* TrimStr - Удаление из строки начальных и конечных пробелов              */
/***************************************************************************/
wchar_t* Far::TrimStr(
  wchar_t *s
  )
{
  // Удаление из строки начальных и конечных пробелов
  return g_FSF.Trim(s);
}
/***************************************************************************/
/* UnquoteStr - Удаление из строки всех двойных кавычек                    */
/***************************************************************************/
wchar_t* Far::UnquoteStr(
  wchar_t *s
  )
{
  // Удаление из строки всех двойных кавычек
  g_FSF.Unquote(s);
  return s;
}
//---------------------------------------------------------------------------
