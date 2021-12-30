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
/* Init - ������������� ���������� Far                                     */
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
/* CreateSettingsObj - �������� ������� "���������" �������                */
/***************************************************************************/
HANDLE Far::CreateSettingsObj()
{
  // �������� ������� "���������"
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
/* FreeSettingsObj - �������� ������� "���������" �������                */
/***************************************************************************/
void Far::FreeSettingsObj(
  HANDLE hSettings
  )
{
  // �������� ������� "���������"
  g_FarSInfo.SettingsControl(hSettings, SCTL_FREE, 0, 0);
}
/***************************************************************************/
/* LoadSetting - �������� �������� ��������� �������                       */
/***************************************************************************/
unsigned __int64 Far::LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  unsigned __int64 nDefault
  )
{
  // �������� �������� �������� �������
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
/* LoadSetting - �������� �������� ��������� �������                       */
/***************************************************************************/
const wchar_t* Far::LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const wchar_t *pwszDefault
  )
{
  // �������� �������� �������� �������
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
/* LoadSetting - �������� �������� ��������� �������                       */
/***************************************************************************/
size_t Far::LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  void *pValue,
  size_t nSize
  )
{
  // �������� �������� �������� �������
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
/* SaveSetting - ���������� �������� ��������� �������                     */
/***************************************************************************/
bool Far::SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  unsigned __int64 nValue
  )
{
  // ���������� �������� �������� �������
  FarSettingsItem item;
  item.StructSize = sizeof(FarSettingsItem);
  item.Root = Root;
  item.Name = pwszName;
  item.Type = FST_QWORD;
  item.Number = nValue;
  return (0 != g_FarSInfo.SettingsControl(hSettings, SCTL_SET, 0, &item));
}
/***************************************************************************/
/* SaveSetting - ���������� �������� ��������� �������                     */
/***************************************************************************/
bool Far::SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const wchar_t *pwszValue
  )
{
  // ���������� �������� �������� �������
  FarSettingsItem item;
  item.StructSize = sizeof(FarSettingsItem);
  item.Root = Root;
  item.Name = pwszName;
  item.Type = FST_STRING;
  item.String = pwszValue;
  return (0 != g_FarSInfo.SettingsControl(hSettings, SCTL_SET, 0, &item));
}
/***************************************************************************/
/* SaveSetting - ���������� �������� ��������� �������                     */
/***************************************************************************/
bool Far::SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const void *pValue,
  size_t nSize
  )
{
  // ���������� �������� �������� �������
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
/* GetLocMsg - ��������� ������ ��������� �� ��������� �����               */
/***************************************************************************/
const wchar_t* Far::GetLocMsg(
  int msgId
  )
{
  return g_FarSInfo.GetMsg(&OLE2VIEWER_GUID, msgId);
}
/***************************************************************************/
/* ShowMessage - ����������� ���������                                     */
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
/* ShowMessage - ����������� ���������                                     */
/***************************************************************************/
void Far::ShowMessage(
  int nFlags,
  int captionId,
  int textId,
  const wchar_t *pwszErrorItem
  )
{
  // ����������� ���������
  ShowMessage(nFlags, GetLocMsg(captionId), GetLocMsg(textId),
              pwszErrorItem);
}
/***************************************************************************/
/* DisplayMenu - ����������� ����                                          */
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
/* ShowConfigDialog - ����������� ����������� ���� ������������ �������    */
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
/* ShowConfirmExtractDialog - ����������� ����������� ���� �������������   */
/*                            ���������� ������                            */
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
/* UpdatePanel - ���������� ���������� ������                              */
/***************************************************************************/
void Far::UpdatePanel()
{
  // ���������� ���������� ������
  g_FarSInfo.PanelControl(PANEL_ACTIVE, FCTL_UPDATEPANEL, 0, NULL);
}
/***************************************************************************/
/* RedrawPanel - ����������� ������                                        */
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
  // ����������� ������
  g_FarSInfo.PanelControl(PANEL_ACTIVE, FCTL_REDRAWPANEL, 0, &ri);
}
/***************************************************************************/
/* SaveScreen - ���������� ������                                          */
/***************************************************************************/
HANDLE Far::SaveScreen()
{
  // ���������� ������� ������
  return g_FarSInfo.SaveScreen(0, 0, -1, -1);
}
/***************************************************************************/
/* RestoreScreen - �������������� ������� ������                           */
/***************************************************************************/
void Far::RestoreScreen(
  HANDLE hScreen
  )
{
  // �������������� ������� ������
  g_FarSInfo.RestoreScreen(hScreen);
}
/***************************************************************************/
/* SetProgressState - ��������� ���� � ��������� ���������� ���������      */
/***************************************************************************/
void Far::SetProgressState(
  int nState
  )
{
  // ��������� ���� � ��������� ���������� ���������
  g_FarSInfo.AdvControl(&OLE2VIEWER_GUID, ACTL_SETPROGRESSSTATE,
                        (intptr_t)nState, NULL);
}
/***************************************************************************/
/* GetPanelInfo - ��������� ����� ���������� �� �������� ������            */
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
/* GetPanelDirectory - ��������� �������� �������� �������� ������         */
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
/* GetCurrentPanelItem - ��������� �������� ��������� �������� ��������    */
/*                       ������                                            */
/***************************************************************************/
size_t Far::GetCurrentPanelItem(
  struct FarGetPluginPanelItem *pItem
  )
{
  // ��������� �������� ��������� �������� ������
  return (size_t)g_FarSInfo.PanelControl(PANEL_ACTIVE,
                                         FCTL_GETCURRENTPANELITEM,
                                         0, pItem);
}
/***************************************************************************/
/* GetCurrentPanelItemFilePath - ��������� ���� � ����� ��� ��������       */
/*                               ��������� �������� �������� ������        */
/*                               (������������ ��������� ����������        */
/*                                ���������� ��� ������ ������� free)      */
/***************************************************************************/
wchar_t* Far::GetCurrentPanelItemFilePath()
{
  struct PanelInfo panelInfo;

  // ��������� ����� ���������� �� �������� ������
  if (!GetPanelInfo(&panelInfo) ||
      (panelInfo.SelectedItemsNumber != 1) ||
      (panelInfo.PanelType != PTYPE_FILEPANEL))
    return NULL;

  size_t nBufSize;

  // ����������� ������� ������ ��� ��������� �������� ��������� ��������
  // ������
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

  // ��������� �������� ��������� �������� ������
  Far::GetCurrentPanelItem(&item);

  wchar_t *pwszFilePath = NULL;

  if (!(pPanelItem->FileAttributes & FILE_ATTRIBUTE_DIRECTORY))
  {
    // ����������� ������� ������ ��� �������� �������� �������� ������
    nBufSize = GetPanelDirectory(NULL, 0);
    if (nBufSize != 0)
    {
      struct FarPanelDirectory *pPanelDir;
      pPanelDir = (struct FarPanelDirectory *)malloc(nBufSize);
      if (pPanelDir)
      {
        // ��������� �������� �������� �������� ������
        pPanelDir->StructSize = sizeof(struct FarPanelDirectory);
        GetPanelDirectory(pPanelDir, nBufSize);

        if (pPanelDir->Name && pPanelDir->Name[0] &&
            IsRelativePath(pPanelItem->FileName))
        {
          // ����������� �����
          pwszFilePath = AllocAndCombinePath(pPanelDir->Name,
                                             pPanelItem->FileName);
        }
        else
        {
          // ����������� ����� �����
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
/* ExpandPath - ��������� ������� ����                                     */
/*              (������������ ��������� ���������� ���������� ��� ������   */
/*               ������� free)                                             */
/***************************************************************************/
wchar_t* Far::ExpandPath(
  const wchar_t *pwszPath
  )
{
  // ����������� ������� ������ ��� ������� ����
  size_t nFullPathSize = g_FSF.ConvertPath(CPM_FULL, pwszPath, NULL, 0);
  if (nFullPathSize == 0)
    return NULL;

  wchar_t *pwszFullPath = (wchar_t *)malloc(nFullPathSize * sizeof(wchar_t));
  if (!pwszFullPath)
    return NULL;

  // ��������� ������� ����
  g_FSF.ConvertPath(CPM_FULL, pwszPath, pwszFullPath, nFullPathSize);

  return pwszFullPath;
}
/***************************************************************************/
/* TrimStr - �������� �� ������ ��������� � �������� ��������              */
/***************************************************************************/
wchar_t* Far::TrimStr(
  wchar_t *s
  )
{
  // �������� �� ������ ��������� � �������� ��������
  return g_FSF.Trim(s);
}
/***************************************************************************/
/* UnquoteStr - �������� �� ������ ���� ������� �������                    */
/***************************************************************************/
wchar_t* Far::UnquoteStr(
  wchar_t *s
  )
{
  // �������� �� ������ ���� ������� �������
  g_FSF.Unquote(s);
  return s;
}
//---------------------------------------------------------------------------
