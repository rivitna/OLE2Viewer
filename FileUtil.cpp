//---------------------------------------------------------------------------
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#include <malloc.h>
#include <memory.h>
#include "StrUtils.h"
#include "FileUtil.h"
//---------------------------------------------------------------------------
#ifndef countof
#define countof(a)  (sizeof(a)/sizeof(a[0]))
#endif  // countof
//---------------------------------------------------------------------------
/***************************************************************************/
/* IsRelativePath - Проверка, является ли указанный путь относительным     */
/***************************************************************************/
bool IsRelativePath(
  const TCHAR *pszPath
  )
{
  if (!pszPath || !pszPath[0])
    return true;
  if (ISPATHDELIMITER(*pszPath))
    return false;
  TCHAR ch = *pszPath | _T(' ');
  return ((ch < _T('a')) || (ch > _T('z')) || (pszPath[1] != _T(':')));
}
/***************************************************************************/
/* IsFileExists - Проверка существования файла                             */
/***************************************************************************/
bool IsFileExists(
  const TCHAR *pszFilePath
  )
{
  return (!(::GetFileAttributes(pszFilePath) & FILE_ATTRIBUTE_DIRECTORY));
}
/***************************************************************************/
/* IsDirectoryExists - Проверка существования каталога                     */
/***************************************************************************/
bool IsDirectoryExists(
  const TCHAR *pszDirPath
  )
{
  DWORD dwAttributes = ::GetFileAttributes(pszDirPath);
  if (dwAttributes == INVALID_FILE_ATTRIBUTES)
    return false;
  return ((dwAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0);
}
/***************************************************************************/
/* ForceDirectories - Создание иерархии каталогов                          */
/***************************************************************************/
bool ForceDirectories(
  const TCHAR *pszDirPath
  )
{
  if (!pszDirPath || !pszDirPath[0])
    return false;

  TCHAR szPath[MAX_PATH];
  void *pPathBuffer = NULL;
  TCHAR *pPath = szPath;

  size_t cchPath = ::lstrlen(pszDirPath) + 1;
  if (cchPath > countof(szPath))
  {
    pPathBuffer = malloc(cchPath * sizeof(TCHAR));
    if (!pPathBuffer)
      return ::SetLastError(ERROR_NOT_ENOUGH_MEMORY), false;
    pPath = (TCHAR *)pPathBuffer;
  }
  memcpy(pPath, pszDirPath, cchPath * sizeof(TCHAR));

  bool bError = false;

  TCHAR *p = pPath;
  if (ISPATHDELIMITER(p[0]) && (p[1] == p[0]))
  {
    p += 2;
    for (int i = 0; i < 2; i++)
    {
      p = StrCharSet(p, _T("\\/"));
      if (!p)
        break;
      p++;
    }
  }
  else
  {
    if ((p[0] != _T('\0')) && (p[1] == _T(':')) && ISPATHDELIMITER(p[2]))
      p += 3;
  }
  while (p && p[0])
  {
    TCHAR chDelim;
    p = StrCharSet(p, _T("\\/"));
    if (p)
    {
      chDelim = *p;
      *p = _T('\0');
    }
    // Создание каталога
    if (!IsDirectoryExists(pPath) && !::CreateDirectory(pPath, NULL))
    {
      bError = true;
      break;
    }
    if (p)
    {
      *p = chDelim;
      p++;
    }
  }

  if (pPathBuffer)
  {
    DWORD dwError;
    if (bError) dwError = ::GetLastError();
    free(pPathBuffer);
    if (bError) ::SetLastError(dwError);
  }
  return !bError;
}
/***************************************************************************/
/* GetFileName - Получение имени файла                                     */
/***************************************************************************/
TCHAR *GetFileName(
  const TCHAR *pszFilePath
  )
{
  if (!pszFilePath)
    return NULL;
  TCHAR *pDelim = StrRCharSet(pszFilePath, _T(":\\/"));
  if (!pDelim)
    return (TCHAR *)pszFilePath;
  return (pDelim + 1);
}
/***************************************************************************/
/* ExtractFilePath - Извлечение пути к файлу                               */
/***************************************************************************/
size_t ExtractFilePath(
  const TCHAR *pszFileName,
  TCHAR *pBuffer,
  size_t nSize
  )
{
  if (!pszFileName)
    return 0;
  TCHAR *pDelim = StrRCharSet(pszFileName, _T(":\\/"));
  if (!pDelim)
    return 0;
  // Перемещение строки в буфер
  return StrCchMoveN(pBuffer, nSize, pszFileName, pDelim - pszFileName + 1);
}
/***************************************************************************/
/* GetFileExt - Получение расширения в имени файла                         */
/***************************************************************************/
TCHAR *GetFileExt(
  const TCHAR *pszFileName
  )
{
  if (!pszFileName)
    return NULL;
  TCHAR *pExt = StrRCharSet(pszFileName, _T(":\\/."));
  if (pExt && pExt[0] == _T('.'))
    return pExt;
  return NULL;
}
/***************************************************************************/
/* AddFileExt - Добавление расширения файла                                */
/***************************************************************************/
size_t AddFileExt(
  const TCHAR *pszFileName,
  const TCHAR *pszExt,
  TCHAR *pBuffer,
  size_t nSize
  )
{
  if (!pszExt || !pszExt[0] ||
      !GetFileExt(pszFileName))
    return 0;
  size_t cchFileName = ::lstrlen(pszFileName);
  size_t cchExt = ::lstrlen(pszExt);
  size_t cchRet = cchFileName + cchExt;
  if (!pBuffer || (nSize == 0))
    return (cchRet + 1);
  memmove(pBuffer, pszFileName, min(cchFileName, nSize - 1) * sizeof(TCHAR));
  if (cchFileName < nSize - 1)
  {
    memcpy(pBuffer + cchFileName, pszExt,
           min(cchExt, nSize - (cchFileName + 1)) * sizeof(TCHAR));
  }
  pBuffer[min(cchRet, nSize - 1)] = _T('\0');
  if (cchRet < nSize)
    return cchRet;
  return nSize;
}
/***************************************************************************/
/* ChangeFileExt - Изменение расширения файла                              */
/***************************************************************************/
size_t ChangeFileExt(
  const TCHAR *pszFileName,
  const TCHAR *pszExt,
  TCHAR *pBuffer,
  size_t nSize
  )
{
  TCHAR *pExt = GetFileExt(pszFileName);
  size_t cchFileName = pExt ? (pExt - pszFileName) : ::lstrlen(pszFileName);
  size_t cchExt = ::lstrlen(pszExt);
  size_t cchRet = cchFileName + cchExt;
  if (!pBuffer || (nSize == 0))
    return (cchRet + 1);
  memmove(pBuffer, pszFileName, min(cchFileName, nSize - 1) * sizeof(TCHAR));
  if (cchFileName < nSize - 1)
  {
    memcpy(pBuffer + cchFileName, pszExt,
           min(cchExt, nSize - (cchFileName + 1)) * sizeof(TCHAR));
  }
  pBuffer[min(cchRet, nSize - 1)] = _T('\0');
  if (cchRet < nSize)
    return cchRet;
  return nSize;
}
/***************************************************************************/
/* InternalCombinePath - Объединение путей                                 */
/***************************************************************************/
size_t InternalCombinePath(
  const TCHAR *pszDir,
  size_t cchDir,
  const TCHAR *pszFile,
  TCHAR *pBuffer,
  size_t nSize
  )
{
  // Указанный путь - относительный?
  if (!IsRelativePath(pszFile))
    return 0;
  size_t cchFile = ::lstrlen(pszFile);
  size_t cchPath = cchDir + cchFile;
  bool bAddDelimiter = ((cchDir != 0) && (cchFile != 0) &&
                        !ISPATHDELIMITER(pszDir[cchDir - 1]));
  if (bAddDelimiter) cchPath++;
  if (!pBuffer || (nSize == 0))
    return (cchPath + 1);
  bool bDirInBuffer = ((pszDir >= pBuffer) && (pszDir < pBuffer + nSize));
  if (bDirInBuffer)
    memmove(pBuffer, pszDir, min(cchDir, nSize - 1) * sizeof(TCHAR));
  if (cchDir < nSize - 1)
  {
    size_t nFileOfs = cchDir;
    if (bAddDelimiter) nFileOfs++;
    memmove(pBuffer + nFileOfs, pszFile,
            min(cchFile, nSize - (nFileOfs + 1)) * sizeof(TCHAR));
    if (bAddDelimiter) pBuffer[cchDir] = _T('\\');
  }
  if (!bDirInBuffer)
    memcpy(pBuffer, pszDir, min(cchDir, nSize - 1) * sizeof(TCHAR));
  pBuffer[min(cchPath, nSize - 1)] = _T('\0');
  if (cchPath < nSize)
    return cchPath;
  return nSize;
}
/***************************************************************************/
/* CombinePath - Объединение путей                                         */
/***************************************************************************/
size_t CombinePath(
  const TCHAR *pszDir,
  const TCHAR *pszFile,
  TCHAR *pBuffer,
  size_t nSize
  )
{
  // Объединение путей
  return InternalCombinePath(pszDir, ::lstrlen(pszDir), pszFile,
                             pBuffer, nSize);
}
/***************************************************************************/
/* AllocAndCombinePath - Выделение памяти и объединение путей              */
/*                       (Возвращенный указатель необходимо освободить     */
/*                        при помощи функции free)                         */
/***************************************************************************/
TCHAR *AllocAndCombinePath(
  const TCHAR *pszDir,
  const TCHAR *pszFile
  )
{
  // Выделение буфера для объединенного пути
  size_t cchAlloc = CombinePath(pszDir, pszFile, NULL, 0);
  if (cchAlloc == 0)
    return NULL;
  TCHAR *pPath = (TCHAR *)malloc(cchAlloc * sizeof(TCHAR));
  if (!pPath)
    return ::SetLastError(ERROR_NOT_ENOUGH_MEMORY), NULL;
  // Объединение путей
  CombinePath(pszDir, pszFile, pPath, cchAlloc);
  return pPath;
}
/***************************************************************************/
/* ChangeFileName - Изменение имени файла                                  */
/***************************************************************************/
size_t ChangeFileName(
  const TCHAR *pszFilePath,
  const TCHAR *pszFile,
  TCHAR *pBuffer,
  size_t nSize
  )
{
  TCHAR *pFile = GetFileName(pszFilePath);
  if (!pFile)
    return 0;
  // Объединение путей
  return InternalCombinePath(pszFilePath, pFile - pszFilePath, pszFile,
                             pBuffer, nSize);
}
//---------------------------------------------------------------------------
