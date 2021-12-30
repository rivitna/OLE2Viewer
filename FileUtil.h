//---------------------------------------------------------------------------
#ifndef __FILEUTIL_H__
#define __FILEUTIL_H__
//---------------------------------------------------------------------------
#ifdef __cplusplus
extern "C" {
#endif  // __cplusplus
//---------------------------------------------------------------------------
#include <stddef.h>
#include <tchar.h>
//---------------------------------------------------------------------------
#define ISPATHDELIMITER(c)  ((c == _T('\\')) || (c == _T('/')))
//---------------------------------------------------------------------------
// ��������, �������� �� ��������� ���� �������������
bool IsRelativePath(
  const TCHAR *pszPath
  );
// �������� ������������� �����
bool IsFileExists(
  const TCHAR *pszFilePath
  );
// �������� ������������� ��������
bool IsDirectoryExists(
  const TCHAR *pszDirPath
  );
// �������� �������� ���������
bool ForceDirectories(
  const TCHAR *pszDirPath
  );
// ��������� ����� �����
TCHAR *GetFileName(
  const TCHAR *pszFilePath
  );
// ���������� ���� � �����
size_t ExtractFilePath(
  const TCHAR *pszFileName,
  TCHAR *pBuffer,
  size_t nSize
  );
// ��������� ���������� � ����� �����
TCHAR *GetFileExt(
  const TCHAR *pszFilePath
  );
// ���������� ���������� � ����� �����
size_t AddFileExt(
  const TCHAR *pszFileName,
  const TCHAR *pszExt,
  TCHAR *pBuffer,
  size_t nSize
  );
// ��������� ���������� � ����� �����
size_t ChangeFileExt(
  const TCHAR *pszFileName,
  const TCHAR *pszExt,
  TCHAR *pBuffer,
  size_t nSize
  );
// ����������� �����
size_t CombinePath(
  const TCHAR *pszDir,
  const TCHAR *pszFile,
  TCHAR *pBuffer,
  size_t nSize
  );
// ����������� �����
// (������������ ��������� ���������� ���������� ��� ������ ������� free)
TCHAR *AllocAndCombinePath(
  const TCHAR *pszDir,
  const TCHAR *pszFile
  );
// ��������� ����� �����
size_t ChangeFileName(
  const TCHAR *pszFilePath,
  const TCHAR *pszFile,
  TCHAR *pBuffer,
  size_t nSize
  );
//---------------------------------------------------------------------------
#ifdef __cplusplus
}  // extern "C"
#endif // __cplusplus
//---------------------------------------------------------------------------
#endif  // __FILEUTIL_H__
