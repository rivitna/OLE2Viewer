//---------------------------------------------------------------------------
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include <stddef.h>
#include <malloc.h>
#include <memory.h>
#include <stdio.h>
#include "Far\plugin.hpp"
#include "Far\DlgBuilder.hpp"
#include <InitGuid.h>
#include "Guids.h"
#include "StrUtils.h"
#include "FileUtil.h"
#include "OLE2Type.h"
#include "OLE2File.h"
#include "FarUtils.h"
#include "StorageObj.h"
#include "PlugLang.h"
#include "PlugInfo.h"
//---------------------------------------------------------------------------
#ifndef _DEBUG
#pragma comment(linker, "/MERGE:.rdata=.text")
#endif  // _DEBUG
//---------------------------------------------------------------------------
// ������ ������ ��� ��������� ������
#define PANEL_TITLE_BUFFER_SIZE  1024
// ����������� ����� � ��������� ������
#define PANEL_TITLE_DELIMITER    L':'

// ������������ ����� ������� �� ����������
#define MAX_EXTENSIONFILTER_LENGTH  4096
// ����������� ���������� � �������
#define EXTENSIONFILTER_DELIMITER   L';'
//---------------------------------------------------------------------------
// ����� ���������� ������������ �������
const wchar_t SETTING_NAME_ENABLED[] = L"Enabled";
const wchar_t SETTING_NAME_USEEXTENSIONFILTER[] = L"UseExtensionFilter";

// ����� ������ � ���������� Ini-����� �������
const wchar_t INI_SECTION_NAME_GENERAL[] = L"General";
const wchar_t INI_KEY_NAME_EXTENSIONFILTER[] = L"ExtensionFilter";

const wchar_t COMMAND_PREFIX[] = L"ole2";

// ������������ ���� ����������
const wchar_t STORAGE_TYPE[] = L"OLE2";
//---------------------------------------------------------------------------
// ��������� ������������ �������
int g_nEnablePlugin;
int g_nUseExtensionFilter;
wchar_t *g_pUpperExtensionFilter;
//---------------------------------------------------------------------------
HMODULE g_hDllHandle;
//---------------------------------------------------------------------------
// ������ �����-����������
int AnalyzeStorage(
  const wchar_t *pwszFileName,
  bool bUseExtensionFilter,
  const wchar_t *pwszUpperExtensionFilter,
  const void *buf,
  size_t bufSize
  );
// �������� �����-����������
HANDLE OpenStorage(
  const wchar_t *pwszFileName
  );
// �������� �����-����������
void CloseStorage(
  HANDLE hStorage
  );
// ��������� ��������� ������
void MakePanelTitle(
  const wchar_t *pwszPluginName,
  const wchar_t *pwszStorageName,
  const wchar_t *pwszCurDir,
  wchar_t *pBuffer,
  size_t nBufferSize
  );
// �������� �������� ������� �� ���������� � ������� ��������
// (������������ ��������� ���������� ���������� ��� ������ ������� free)
wchar_t *LoadUpperExtensionFilter();
// �������� ������������ ������� �� ����������
bool IsFileExtensionFilterMatch(
  const wchar_t *pwszFileName,
  const wchar_t *pwszUpperExtensionFilter
  );
//---------------------------------------------------------------------------
/***************************************************************************/
/* DllMain                                                                 */
/***************************************************************************/
BOOL WINAPI DllMain(
  HINSTANCE hinstDLL,
  DWORD fdwReason,
  LPVOID lpvReserved
  )
{
  switch (fdwReason)
  {
    case DLL_PROCESS_ATTACH:
      g_hDllHandle = hinstDLL;
      break;
  }
  return TRUE;
}
//---------------------------------------------------------------------------
/***************************************************************************/
/* GetGlobalInfoW - ��������� �������� ���������� � �������                */
/***************************************************************************/
void WINAPI GetGlobalInfoW(
  struct GlobalInfo *Info
  )
{
  Info->StructSize = sizeof(struct GlobalInfo);
  Info->MinFarVersion = MAKEFARVERSION(MIN_FAR_VERSION_MAJOR,
                                       MIN_FAR_VERSION_MINOR,
                                       MIN_FAR_VERSION_REVISION,
                                       MIN_FAR_VERSION_BUILD,
                                       VS_RELEASE);
  Info->Version = MAKEFARVERSION(OLE2VIEWER_VERSION_MAJOR,
                                 OLE2VIEWER_VERSION_MINOR,
                                 OLE2VIEWER_VERSION_REVISION,
                                 OLE2VIEWER_VERSION_BUILD,
                                 OLE2VIEWER_VERSION_STAGE);
  Info->Guid = OLE2VIEWER_GUID;
  Info->Title = OLE2VIEWER_TITLE;
  Info->Description = OLE2VIEWER_DESCRIPTION;
  Info->Author = OLE2VIEWER_AUTHOR;
}
/***************************************************************************/
/* SetStartupInfoW - �������� ������� ����������� ����������               */
/***************************************************************************/
void WINAPI SetStartupInfoW(
  const struct PluginStartupInfo *Info
  )
{
  // ������������� ���������� Far
  Far::Init(Info);

  // �������� ���������� ������������ �������
  g_nEnablePlugin = 1;
  g_nUseExtensionFilter = 0;
  g_pUpperExtensionFilter = NULL;
  // �������� ������� "���������" �������
  HANDLE hSettings = Far::CreateSettingsObj();
  if (hSettings != INVALID_HANDLE_VALUE)
  {
    // �������� ��������� ��������� �������
    g_nEnablePlugin = (int)Far::LoadSetting(hSettings, 0,
                                            SETTING_NAME_ENABLED,
                                            g_nEnablePlugin);
    // �������� ��������� ��������� ������� �� ����������
    g_nUseExtensionFilter =
      (int)Far::LoadSetting(hSettings, 0, SETTING_NAME_USEEXTENSIONFILTER,
                            g_nUseExtensionFilter);
    // �������� ������� "���������" �������
    Far::FreeSettingsObj(hSettings);
  }
  if (g_nUseExtensionFilter)
  {
    //�������� �������� ������� �� ���������� � ������� ��������
    g_pUpperExtensionFilter = LoadUpperExtensionFilter();
  }
}
/***************************************************************************/
/* GetPluginInfoW - ��������� �������������� ���������� � �������          */
/***************************************************************************/
void WINAPI GetPluginInfoW(
  struct PluginInfo *Info
  )
{
  Info->StructSize = sizeof(struct PluginInfo);
  Info->Flags = 0;

  static const wchar_t *pluginMenuStrings[1];
  pluginMenuStrings[0] = Far::GetLocMsg(MSG_PLUGIN_NAME);

  static const wchar_t *configMenuStrings[1];
  configMenuStrings[0] = Far::GetLocMsg(MSG_PLUGIN_NAME);

  Info->PluginMenu.Guids = &GUID_O2V_INFO_MENU;
  Info->PluginMenu.Strings = pluginMenuStrings;
  Info->PluginMenu.Count = countof(pluginMenuStrings);
  Info->PluginConfig.Guids = &GUID_O2V_INFO_CONFIG;
  Info->PluginConfig.Strings = configMenuStrings;
  Info->PluginConfig.Count = countof(configMenuStrings);
  Info->CommandPrefix = COMMAND_PREFIX;
}
/***************************************************************************/
/* ExitFARW - ����� �� Far Manager                                         */
/***************************************************************************/
void WINAPI ExitFARW(
  const ExitInfo *Info
  )
{
  // ������������ ������, ����������� ��� ������� �� ����������
  if (g_pUpperExtensionFilter)
  {
    free(g_pUpperExtensionFilter);
    g_pUpperExtensionFilter = NULL;
  }
}
/***************************************************************************/
/* ConfigureW - ��������� ������������ �������                             */
/***************************************************************************/
intptr_t WINAPI ConfigureW(
  const struct ConfigureInfo *Info
  )
{
  if (Info->StructSize < sizeof(struct ConfigureInfo))
    return 0;  /* return nullptr; */

  // ����������� ����������� ���� ������������ �������
  int nEnablePlugin = g_nEnablePlugin;
  int nUseExtensionFilter = g_nUseExtensionFilter;
  if (Far::ShowConfigDialog(&nEnablePlugin,
                            &nUseExtensionFilter))
  {
    g_nEnablePlugin = nEnablePlugin;
    if (nUseExtensionFilter != g_nUseExtensionFilter)
    {
      g_nUseExtensionFilter = nUseExtensionFilter;
      // ������������ ������, ����������� ��� ������� �� ����������
      if (g_pUpperExtensionFilter)
      {
        free(g_pUpperExtensionFilter);
        g_pUpperExtensionFilter = NULL;
      }
      if (nUseExtensionFilter)
      {
        //�������� �������� ������� �� ���������� � ������� ��������
        g_pUpperExtensionFilter = LoadUpperExtensionFilter();
      }
    }

    // ���������� ���������� ������������ �������
    // �������� ������� "���������" �������
    HANDLE hSettings = Far::CreateSettingsObj();
    if (hSettings != INVALID_HANDLE_VALUE)
    {
      // ���������� ��������� ��������� �������
      Far::SaveSetting(hSettings, 0, SETTING_NAME_ENABLED, nEnablePlugin);
      // ���������� ��������� ��������� ������� �� ����������
      Far::SaveSetting(hSettings, 0, SETTING_NAME_USEEXTENSIONFILTER,
                       nUseExtensionFilter);
      // �������� ������� "���������" �������
      Far::FreeSettingsObj(hSettings);
    }
  }
  return 0;
}
/***************************************************************************/
/* AnalyseW - ������ ����������� �����                                     */
/***************************************************************************/
HANDLE WINAPI AnalyseW(
  const struct AnalyseInfo *Info
  )
{
  if (!Info ||
      (Info->StructSize < sizeof(struct AnalyseInfo)) ||
      !g_nEnablePlugin ||
      !Info->FileName)
    return (HANDLE)0;  /* return nullptr; */
  // ������ �����-����������
  return (HANDLE)AnalyzeStorage(Info->FileName,
                                (g_nUseExtensionFilter != 0),
                                g_pUpperExtensionFilter,
                                Info->Buffer,
                                Info->BufferSize);
}
/***************************************************************************/
/* CloseAnalyseW - ������������ ��������                                   */
/***************************************************************************/
void WINAPI CloseAnalyseW(
  const struct AnalyseInfo *Info
  )
{
}
/***************************************************************************/
/* OpenW - ������ �������                                                  */
/***************************************************************************/
HANDLE WINAPI OpenW(
  const struct OpenInfo *Info
  )
{
  if (Info->StructSize < sizeof(struct OpenInfo))
    return (HANDLE)0;  /* return nullptr; */

  bool bDoOpen = true;
  const wchar_t *pwszFilePath = NULL;
  wchar_t *pwszCmdLine = NULL;

  switch (Info->OpenFrom)
  {
    case OPEN_PLUGINSMENU:
      // �������� ������� �� ���� "������� ��������"
      // ��������� ���� � ����� ��� �������� ��������� �������� ��������
      // ������
      pwszFilePath = Far::GetCurrentPanelItemFilePath();
      if (pwszFilePath)
      {
        // ����������� ����
        struct FarMenuItem menuItems[1];
        memset(menuItems, 0, sizeof(menuItems));
        menuItems[0].Text = Far::GetLocMsg(MSG_OPEN_MENU_OPEN);
        intptr_t nMenuItem;
        nMenuItem = Far::DisplayMenu(Far::GetLocMsg(MSG_PLUGIN_NAME),
                                     menuItems, countof(menuItems));
        if (nMenuItem != 0)
          bDoOpen = false;
      }
      break;

    case OPEN_COMMANDLINE:
      // �������� ������� �� ��������� ������
      // ����������� ��������� ������
      pwszCmdLine =
        StrDup(SkipBlanksW(((const struct OpenCommandLineInfo *)Info->Data)->CommandLine));
      if (pwszCmdLine && pwszCmdLine[0])
      {
        // �������� �� ������ ���� ������� �������
        Far::UnquoteStr(pwszCmdLine);
        // ��������� ������� ���� � �����
        pwszFilePath = Far::ExpandPath(pwszCmdLine);
      }
      break;

    case OPEN_ANALYSE:
      // �������� ������� ����� ������� ����������� �����
      pwszFilePath =
        ((const struct OpenAnalyseInfo *)Info->Data)->Info->FileName;
      break;
  }

  HANDLE hStorage = (HANDLE)0;

  if (bDoOpen && pwszFilePath && pwszFilePath[0])
  {
    // �������� �����-����������
    hStorage = OpenStorage(pwszFilePath);
  }

  if ((Info->OpenFrom != OPEN_ANALYSE) && pwszFilePath)
  {
    // ����������� ���� � �����
    free((void *)pwszFilePath);
  }

  if (pwszCmdLine)
  {
    // ����������� ��������� ������
    free(pwszCmdLine);
  }

  return hStorage;
}
/***************************************************************************/
/* ClosePanelW - �������� �������� ������ �������                          */
/***************************************************************************/
void WINAPI ClosePanelW(
  const struct ClosePanelInfo *Info
)
{
  if ((Info->StructSize < sizeof(struct ClosePanelInfo)) ||
      !Info->hPanel)
    return;

  // �������� �����-����������
  CloseStorage(Info->hPanel);
}
/***************************************************************************/
/* GetFindDataW - ��������� ������ ����� �� �������� ��������              */
/***************************************************************************/
intptr_t WINAPI GetFindDataW(
  struct GetFindDataInfo *Info
  )
{
  if ((Info->StructSize < sizeof(struct GetFindDataInfo)) ||
      !Info->hPanel)
    return 0;

  CStorageObject *pStorage = (CStorageObject *)Info->hPanel;

  // ��������� ���������� ��������� � ������� ��������
  size_t nItemCount = pStorage->GetItemCount(true);
  if (nItemCount == 0)
    return 1;

  struct PluginPanelItem *pPanelItem =
    (struct PluginPanelItem *)malloc(nItemCount *
                                     sizeof(struct PluginPanelItem));
  if (!pPanelItem)
    return 0;

  Info->PanelItem = pPanelItem;

  memset(pPanelItem, 0, nItemCount * sizeof(struct PluginPanelItem));

  size_t i = 0;

  // ������������ ��������� �������� ��������
  FindItemData findItemData;
  if (pStorage->FindFirstItem(true, &findItemData))
  {
    do
    {
      pPanelItem->FileAttributes = findItemData.dwItemAttributes;
      pPanelItem->CreationTime.dwLowDateTime =
        findItemData.ftCreationTime.dwLowDateTime;
      pPanelItem->CreationTime.dwHighDateTime =
        findItemData.ftCreationTime.dwHighDateTime;
      pPanelItem->LastWriteTime.dwLowDateTime =
        findItemData.ftModifiedTime.dwLowDateTime;
      pPanelItem->LastWriteTime.dwHighDateTime =
        findItemData.ftModifiedTime.dwHighDateTime;
      pPanelItem->ChangeTime.dwLowDateTime =
        findItemData.ftModifiedTime.dwLowDateTime;
      pPanelItem->ChangeTime.dwHighDateTime =
        findItemData.ftModifiedTime.dwHighDateTime;
      pPanelItem->FileSize = findItemData.nItemSize;
      pPanelItem->FileName = findItemData.pwszItemName;
      pPanelItem->UserData.Data = (void *)findItemData.dwItemID;
      i++;
      pPanelItem++;
    }
    while ((i != nItemCount) && pStorage->FindNextItem(&findItemData));
  }

  Info->ItemsNumber = i;

  return 1;
}
/***************************************************************************/
/* FreeFindDataW - ��������� ������ ����� �� �������� ��������              */
/***************************************************************************/
void WINAPI FreeFindDataW(
  const struct FreeFindDataInfo *Info
  )
{
  if (Info->StructSize < sizeof(struct FreeFindDataInfo))
    return;
  free(Info->PanelItem);
}
/***************************************************************************/
/* SetDirectoryW - ��������� �������� �������� �� ����������� ��������     */
/*                 �������                                                 */
/***************************************************************************/
intptr_t WINAPI SetDirectoryW(
  const struct SetDirectoryInfo *Info
  )
{
  if ((Info->StructSize < sizeof(struct SetDirectoryInfo)) ||
      !Info->hPanel)
    return 0;

  if (!Info->Dir || !Info->Dir[0])
    return 1;

  CStorageObject *pStorage = (CStorageObject *)Info->hPanel;

  // ��������� �������� ��������
  return (intptr_t)pStorage->SetCurrentDir(Info->Dir);
}
/***************************************************************************/
/* GetOpenPanelInfoW - ��������� ���������� �� ����������� ������ �������  */
/***************************************************************************/
void WINAPI GetOpenPanelInfoW(
  struct OpenPanelInfo *Info
  )
{
  Info->StructSize = sizeof(struct OpenPanelInfo);

  if (!Info->hPanel)
    return;

  CStorageObject *pStorage = (CStorageObject *)Info->hPanel;

  static wchar_t wszTitle[PANEL_TITLE_BUFFER_SIZE];
  static wchar_t wszVersionInfo[32];
  static wchar_t wszSizeInfo[32];
  static wchar_t wszStreamCount[32];
  static wchar_t wszStorageCount[32];
  static wchar_t wszMacroCount[32];

  // ��������� ����� �������
  const wchar_t *pwszPluginName = Far::GetLocMsg(MSG_PLUGIN_NAME);
  // ��������� ����� �����-����������
  const wchar_t *pwszStorageName = pStorage->GetFileName();
  // ��������� ���� � �������� ��������
  const wchar_t *pwszCurDir = pStorage->GetCurrentDirPath();

  // ��������� ��������� ������
  MakePanelTitle(pwszPluginName, pwszStorageName, pwszCurDir,
                 wszTitle, countof(wszTitle));

  // FAR �� ������� �� �������, ���� �������� ������� - "\"
  if (pwszCurDir && pwszCurDir[0] && !pwszCurDir[1])
    pwszCurDir = L"";

  // ��������� ������ � ������� ����� OLE2
  swprintf_s(wszVersionInfo, countof(wszVersionInfo), L"%d.%d",
             pStorage->GetMajorVersion(), pStorage->GetMinorVersion());
  // ��������� ������ � �������� ����� OLE2
  UI64ToStrW(pStorage->GetFileSize(), L' ', wszSizeInfo);
  // ��������� ������ � ����������� ������� � ����� OLE2
  UI64ToStrW(pStorage->GetStreamCount(), L' ', wszStreamCount);
  // ��������� ������ � ����������� ��������� � ����� OLE2
  UI64ToStrW(pStorage->GetStorageCount(), L' ', wszStorageCount);
  // ��������� ������ � ����������� �������� � ����� OLE2
  UI64ToStrW(pStorage->GetMacroCount(), L' ', wszMacroCount);

  static InfoPanelLine infoLines[8];

  memset(infoLines, 0, sizeof(infoLines));

  infoLines[0].Text = pwszStorageName;
  infoLines[0].Flags = IPLFLAGS_SEPARATOR;

  infoLines[1].Text = Far::GetLocMsg(MSG_INFOL_VERSION);
  infoLines[1].Data = wszVersionInfo;

  infoLines[2].Text = Far::GetLocMsg(MSG_INFOL_TYPE);
  infoLines[2].Data = OLE2FileTypeToString(pStorage->GetFileType());

  infoLines[3].Text = Far::GetLocMsg(MSG_INFOL_SIZE);
  infoLines[3].Data = wszSizeInfo;

  infoLines[4].Flags = IPLFLAGS_SEPARATOR;

  infoLines[5].Text = Far::GetLocMsg(MSG_INFOL_STREAMS);
  infoLines[5].Data = wszStreamCount;

  infoLines[6].Text = Far::GetLocMsg(MSG_INFOL_STORAGES);
  infoLines[6].Data = wszStorageCount;

  infoLines[7].Text = Far::GetLocMsg(MSG_INFOL_MACROS);
  infoLines[7].Data = wszMacroCount;

  // ��������� ������� ������
  static KeyBarLabel key_labels[] =
  {
    { { VK_F6, 0 }, L"", L"" },
    { { VK_F6, SHIFT_PRESSED }, L"", L"" },
    { { VK_F7, 0 }, L"", L"" },
    { { VK_F8, 0 }, L"", L"" },
    { { VK_F8, SHIFT_PRESSED }, L"", L"" }
  };
  static KeyBarTitles kb_titles = { countof(key_labels), key_labels };

  Info->Flags = OPIF_DISABLEFILTER | OPIF_DISABLESORTGROUPS;
  // ��������� ����� �����-����������
  Info->HostFile = pStorage->GetFilePath();
  // ��������� ���� � �������� ��������
  Info->CurDir = pwszCurDir;
  // ��������� ��������� ������ �������
  Info->PanelTitle = wszTitle;
  // ��������� ����� � �������������� ������
  Info->InfoLinesNumber = countof(infoLines);
  Info->InfoLines = infoLines;
  // ��������� �������� ������� ������
  Info->KeyBar = &kb_titles;
}
/***************************************************************************/
/* GetFilesW - ��������� ������ ��� ���������                              */
/***************************************************************************/
intptr_t WINAPI GetFilesW(
  struct GetFilesInfo *Info
  )
{
  if ((Info->StructSize < sizeof(struct GetFilesInfo)) ||
      Info->Move ||
      !Info->DestPath ||
      (Info->ItemsNumber == 0) ||
      !Info->hPanel)
    return 0;

  // ������ ������ ������� ".."
  if ((Info->ItemsNumber == 1) &&
      (Info->PanelItem[0].FileName[0] == L'.') &&
      (Info->PanelItem[0].FileName[1] == L'.') &&
      (Info->PanelItem[0].FileName[2] == L'\0'))
    return 0;

  const wchar_t *pwszDestPath = Info->DestPath;
  wchar_t wszDestPathBuf[MAX_PATH];

  if (!(Info->OpMode & OPM_SILENT))
  {
    StrCchCopy(wszDestPathBuf, countof(wszDestPathBuf), Info->DestPath);

    // ����������� ����������� ���� ������������� ���������� ������
    if (!Far::ShowConfirmExtractDialog(wszDestPathBuf,
                                       countof(wszDestPathBuf)))
      return -1;

    pwszDestPath = wszDestPathBuf;
  }

  CStorageObject *pStorage = (CStorageObject *)Info->hPanel;

  bool bError = false;

  // ���������� ���������
  for (size_t i = 0; i < Info->ItemsNumber; i++)
  {
    // ���������� ��������
    if (!pStorage->ExtractItem((unsigned long)Info->PanelItem[i].UserData.Data,
                               Info->PanelItem[i].FileName,
                               pwszDestPath))
    {
      bError = true;
    }
    else
    {
      Info->PanelItem[i].Flags &= ~PPIF_SELECTED;
    }
  }

  return !bError ? 1 : 0;
}
/***************************************************************************/
/* ProcessPanelInputW - ��������� ������� ���������� � ����                */
/***************************************************************************/
intptr_t WINAPI ProcessPanelInputW(
  const struct ProcessPanelInputInfo *Info
  )
{
  if ((Info->StructSize < sizeof(struct ProcessPanelInputInfo)) ||
      (Info->Rec.EventType != KEY_EVENT) ||
      !Info->hPanel)
    return 0;

  KEY_EVENT_RECORD keyEvtRec = Info->Rec.Event.KeyEvent;
  if (!keyEvtRec.bKeyDown ||
      ((keyEvtRec.wVirtualKeyCode != VK_RETURN) ||
       (keyEvtRec.dwControlKeyState & (LEFT_ALT_PRESSED |
                                       RIGHT_ALT_PRESSED |
                                       SHIFT_PRESSED))) &&
      ((keyEvtRec.wVirtualKeyCode != VK_NEXT) ||
       !(keyEvtRec.dwControlKeyState & (LEFT_CTRL_PRESSED |
                                        RIGHT_CTRL_PRESSED))))
    return 0;

  // ����������� ������� ������ ��� ��������� �������� ��������� ��������
  // ������
  size_t nBufSize = Far::GetCurrentPanelItem(NULL);
  if (nBufSize == 0)
    return 0;

  struct FarGetPluginPanelItem item;
  item.StructSize = sizeof(struct FarGetPluginPanelItem);
  item.Size = nBufSize;
  struct PluginPanelItem *pPanelItem;
  pPanelItem = (struct PluginPanelItem *)malloc(nBufSize);
  if (!pPanelItem)
    return 0;
  item.Item = pPanelItem;

  // ��������� �������� ��������� �������� ������
  Far::GetCurrentPanelItem(&item);

  unsigned long dwStreamID = NOSTREAM;

  if (!(pPanelItem->FileAttributes & FILE_ATTRIBUTE_DIRECTORY))
  {
    // ��������� �������������� ������
    dwStreamID = (unsigned long)pPanelItem->UserData.Data;
  }

  free(pPanelItem);

  if (dwStreamID != NOSTREAM)
  {
    CStorageObject *pStorage = (CStorageObject *)Info->hPanel;
    // �������� ������
    if (pStorage->OpenStream(dwStreamID))
    {
      // ���������� ���������� ������
      Far::UpdatePanel();
      // ����������� ������
      Far::RedrawPanel(0, 0);
      return 1;
    }
  }

  return 0;
}
//---------------------------------------------------------------------------
/***************************************************************************/
/* AnalyzeStorage - ������ �����-����������                                */
/***************************************************************************/
int AnalyzeStorage(
  const wchar_t *pwszFileName,
  bool bUseExtensionFilter,
  const wchar_t *pwszUpperExtensionFilter,
  const void *buf,
  size_t bufSize
  )
{
  if (bUseExtensionFilter)
  {
    // �������� ������������ ������� �� ����������
    if (!IsFileExtensionFilterMatch(pwszFileName, pwszUpperExtensionFilter))
      return 0;
  }

  if ((bufSize < 2 * sizeof(uint32_t)) ||
      (((uint32_t *)buf)[0] != OLE2_SIGNATURE) ||
      (((uint32_t *)buf)[1] != OLE2_VERSION))
    return 0;

  const PCOMPOUND_FILE_HEADER pFileHeader = (PCOMPOUND_FILE_HEADER)buf;
  if (bufSize >= sizeof_through_field(COMPOUND_FILE_HEADER, wSectorShift))
  {
    if ((pFileHeader->wSectorShift != 0x0009) &&
        (pFileHeader->wSectorShift != 0x000C))
      return 0;
    if (bufSize >= sizeof_through_field(COMPOUND_FILE_HEADER,
                                        wMiniSectorShift))
    {
      if (pFileHeader->wMiniSectorShift != 6)
        return 0;
    }
  }
  return 1;
}
/***************************************************************************/
/* OpenStorage - �������� �����-����������                                 */
/***************************************************************************/
HANDLE OpenStorage(
  const wchar_t *pwszFileName
  )
{
  HANDLE hStorage = (HANDLE)0;  /* nullptr; */

  // �������� ������� �����-����������
  CStorageObject *pStorage = new CStorageObject();

  // ���������� ������
  HANDLE hScreen = Far::SaveScreen();

  // ����������� ��������� �� �������� �����-����������
  Far::ShowMessage(0, MSG_PLUGIN_NAME, MSG_OPEN_FILE_OPENING, NULL);

  // ��������� ���� � ��������� ���������� ���������
  Far::SetProgressState(TBPS_INDETERMINATE);

  // �������� �����-����������
  int res = pStorage->Open(pwszFileName);
  if (pStorage->GetOpened())
  {
    hStorage = (HANDLE)pStorage;
    // ������������ �������� ��� ����������� ������� ��� �� ���������
    pStorage->UseOpenableStreamAttr = true;
  }
  else
  {
    // ����������� ������� �����-����������
    delete pStorage;
  }

  if ((res != O2_OK) && (res != O2_ERROR_SIGN))
  {
    // ����������� ��������� �� ������
    int errMsgId;
    switch (res)
    {
      case O2_ERROR_PARAM: errMsgId = MSG_ERROR_PARAM; break;
      case O2_ERROR_DATA: errMsgId = MSG_ERROR_DATA; break;
      case O2_ERROR_MEM: errMsgId = MSG_ERROR_MEM; break;
      case O2_ERROR_OPEN: errMsgId = MSG_ERROR_OPEN; break;
      case O2_ERROR_READ: errMsgId = MSG_ERROR_READ; break;
      case O2_ERROR_WRITE: errMsgId = MSG_ERROR_WRITE; break;
      default: errMsgId = MSG_ERROR_UNKNOWN; break;
    }
    Far::ShowMessage(FMSG_WARNING | FMSG_MB_OK, MSG_PLUGIN_NAME, errMsgId,
                     GetFileName(pwszFileName));
  }

  // ��������� ���� � ��������� ���������� ���������
  Far::SetProgressState(TBPS_NOPROGRESS);

  // �������������� ������
  Far::RestoreScreen(hScreen);

  return hStorage;
}
/***************************************************************************/
/* CloseStorage - �������� �����-����������                                */
/***************************************************************************/
void CloseStorage(
  HANDLE hStorage
  )
{
  CStorageObject *pStorage = (CStorageObject *)hStorage;
  // �������� �����-����������
  pStorage->Close();
  delete pStorage;
}
/***************************************************************************/
/* MakePanelTitle - ��������� ��������� ������                             */
/***************************************************************************/
void MakePanelTitle(
  const wchar_t *pwszPluginName,
  const wchar_t *pwszStorageName,
  const wchar_t *pwszCurDir,
  wchar_t *pBuffer,
  size_t nBufferSize
  )
{
  struct
  {
    size_t         len;
    const wchar_t *s;
  } strings[4];

  strings[0].s = pwszPluginName;
  strings[1].s = pwszStorageName;
  strings[2].s = STORAGE_TYPE;
  strings[3].s = pwszCurDir;

  size_t cchTitle = 0;
  size_t i;
  wchar_t *p;

  // ����������� ���� ����� � ����� ����� ���������
  for (i = 0; i < countof(strings); i++)
  {
    strings[i].len = ::lstrlen(strings[i].s);
    if (strings[i].len != 0)
      cchTitle += strings[i].len + 1;
  }
  if (cchTitle == 0) cchTitle++;

  if (cchTitle <= nBufferSize)
  {
    // ����������� ���� ����� � �������������
    p = pBuffer;
    for (i = 0; i < countof(strings); i++)
    {
      if (strings[i].len != 0)
      {
        memcpy(p, strings[i].s, strings[i].len * sizeof(wchar_t));
        p += strings[i].len;
        if (i != countof(strings) - 1)
          *(p++) = PANEL_TITLE_DELIMITER;
      }
    }
    *p = L'\0';
  }
  else
  {
    // ����������� ����������� ����� � �����
    nBufferSize--;
    p = pBuffer + nBufferSize;
    *p = L'\0';
    i = countof(strings);
    while ((nBufferSize != 0) && (i != 0))
    {
      i--;
      size_t nToCopy = min(nBufferSize, strings[i].len);
      if (nToCopy != 0)
      {
        p -= nToCopy;
        memcpy(p, strings[i].s + (strings[i].len - nToCopy),
               nToCopy * sizeof(wchar_t));
        nBufferSize -= nToCopy;
        if ((i != 0) && (nBufferSize != 0))
        {
          p--;
          *p = PANEL_TITLE_DELIMITER;
          nBufferSize--;
        }
      }
    }
  }
}
/***************************************************************************/
/* GetPluginIniFileName - ��������� ����� Ini-����� �������                */
/***************************************************************************/
bool GetPluginIniFileName(
  wchar_t *pwszIniFileName,
  size_t nSize
  )
{
  size_t cch;

  // ��������� ���� � ������������ ����� �������
  cch = ::GetModuleFileNameW(g_hDllHandle, pwszIniFileName, (DWORD)nSize);
  if (!cch || (cch == nSize))
    return false;
  // ��������� ���������� � ����� �����
  cch = ChangeFileExt(pwszIniFileName, L".ini", pwszIniFileName, nSize);
  if (!cch || (cch >= nSize))
    return false;
  return true;
}
/***************************************************************************/
/* LoadUpperExtensionFilter - �������� �������� ������� �� ���������� �    */
/*                            ������� ��������                             */
/*                            (������������ ��������� ����������           */
/*                             ���������� ��� ������ ������� free)         */
/***************************************************************************/
wchar_t *LoadUpperExtensionFilter()
{
  wchar_t wszIniFileName[MAX_PATH];

  // ��������� ����� Ini-����� �������
  if (!GetPluginIniFileName(wszIniFileName, countof(wszIniFileName)))
    return NULL;

  wchar_t buf[MAX_EXTENSIONFILTER_LENGTH];

  // ��������� �������� ������� �� ����������
  DWORD dwLen = ::GetPrivateProfileString(INI_SECTION_NAME_GENERAL,
                                          INI_KEY_NAME_EXTENSIONFILTER,
                                          L"", buf, countof(buf),
                                          wszIniFileName);
  if (!dwLen)
    return NULL;

  // �������������� �������� � ������� �������
  ::CharUpperBuff(buf, dwLen);

  return StrDup(buf);
}
/***************************************************************************/
/* IsFileExtensionFilterMatch - �������� ������������ ������� �� ����������*/
/***************************************************************************/
bool IsFileExtensionFilterMatch(
  const wchar_t *pwszFileName,
  const wchar_t *pwszUpperExtensionFilter
  )
{
  if (!pwszFileName || !pwszFileName[0] ||
      !pwszUpperExtensionFilter || !pwszUpperExtensionFilter[0])
    return false;

  const wchar_t *pExt = GetFileExt(pwszFileName);
  if (!pExt || !pExt[0])
    return false;

  wchar_t wszUpperExt[MAX_PATH];

  pExt++;
  size_t cchExt = ::lstrlen(pExt);
  if ((cchExt == 0) || (cchExt > countof(wszUpperExt)))
    return false;

  memcpy(wszUpperExt, pExt, cchExt * sizeof(wchar_t));
  ::CharUpperBuff(wszUpperExt, (DWORD)cchExt);

  const wchar_t *p = pwszUpperExtensionFilter;
  for (;;)
  {
    const wchar_t *pDelim = wcschr(p, EXTENSIONFILTER_DELIMITER);
    size_t cch = pDelim ? pDelim - p : lstrlen(p);
    if ((cch == cchExt) && !memcmp(p, wszUpperExt, cch * sizeof(wchar_t)))
      return true;
    if (!pDelim)
      break;
    p = pDelim + 1;
  }

  return false;
}
//---------------------------------------------------------------------------
