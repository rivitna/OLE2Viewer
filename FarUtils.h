//---------------------------------------------------------------------------
#ifndef __FARUTILS_H__
#define __FARUTILS_H__
//---------------------------------------------------------------------------
namespace Far {
//---------------------------------------------------------------------------
// ������������� ���������� Far
void Init(
  const struct PluginStartupInfo *pInfo
  );
// �������� ������� "���������" �������
HANDLE CreateSettingsObj();
// �������� ������� "���������" �������
void FreeSettingsObj(
  HANDLE hSettings
  );
// �������� �������� ��������� �������
unsigned __int64 LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  unsigned __int64 nDefault
  );
// �������� �������� ��������� �������
const wchar_t* LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const wchar_t *pwszDefault
  );
// �������� �������� ��������� �������
size_t LoadSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  void *pValue,
  size_t nSize
  );
// ���������� �������� ��������� �������
bool SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  unsigned __int64 nValue
  );
// ���������� �������� ��������� �������
bool SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const wchar_t *pwszValue
  );
// ���������� �������� ��������� �������
bool SaveSetting(
  HANDLE hSettings,
  size_t Root,
  const wchar_t *pwszName,
  const void *pValue,
  size_t nSize
  );
// ��������� ������ ��������� �� ��������� �����
const wchar_t* GetLocMsg(
  int msgId
  );
// ����������� ���������
void ShowMessage(
  int nFlags,
  const wchar_t *pwszCaption,
  const wchar_t *pwszText,
  const wchar_t *pwszErrorItem
  );
// ����������� ���������
void ShowMessage(
  int nFlags,
  int captionId,
  int textId,
  const wchar_t *pwszErrorItem
  );
// ����������� ����
intptr_t DisplayMenu(
  const wchar_t *pwszTitle,
  const struct FarMenuItem *pItems,
  size_t nNumItems
  );
// ����������� ����������� ���� ������������ �������
bool ShowConfigDialog(
  int *pnEnablePlugin,
  int *pnUseExtensionFilter
  );
// ����������� ����������� ���� ������������� ���������� ������
bool ShowConfirmExtractDialog(
  wchar_t *pwszDestPathBuf,
  size_t nDestPathBufSize
  );
// ���������� ���������� ������
void UpdatePanel();
// ����������� ������
void RedrawPanel(
  size_t nCurrentItem,
  size_t nTopPanelItem
  );
// ���������� ������
HANDLE SaveScreen();
// �������������� ������� ������
void RestoreScreen(
  HANDLE hScreen
  );
// ��������� ���� � ��������� ���������� ���������
void SetProgressState(
  int nState
  );
// ��������� ����� ���������� �� �������� ������
bool GetPanelInfo(
  struct PanelInfo *pPanelInfo
  );
// ��������� �������� �������� �������� ������
size_t GetPanelDirectory(
  struct FarPanelDirectory *pPanelDir,
  size_t nBufSize
  );
// ��������� �������� ��������� �������� �������� ������
size_t GetCurrentPanelItem(
  struct FarGetPluginPanelItem *pItem
  );
// ��������� ���� � ����� ��� �������� ��������� �������� �������� ������
// (������������ ��������� ���������� ���������� ��� ������ ������� free)
wchar_t* GetCurrentPanelItemFilePath();
// ��������� ������� ����
// (������������ ��������� ���������� ���������� ��� ������ ������� free)
wchar_t* ExpandPath(
  const wchar_t *pwszPath
  );
// �������� �� ������ ��������� � �������� ��������
wchar_t* TrimStr(
  wchar_t *s
  );
// �������� �� ������ ���� ������� �������
wchar_t* UnquoteStr(
  wchar_t *s
  );
//---------------------------------------------------------------------------
}  // namespace Far
//---------------------------------------------------------------------------
#endif  // __FARUTILS_H__
