//---------------------------------------------------------------------------
#ifndef USE_WINAPI
#include <errno.h>
#endif  // USE_WINAPI

#include <tchar.h>
#include "FileStreams.h"
//---------------------------------------------------------------------------
/***************************************************************************/
/* ~CFileInStream - Деструктор класса                                      */
/***************************************************************************/
CFileInStream::~CFileInStream()
{
  // Закрытие
  Close();
}
/***************************************************************************/
/* Open - Открытие                                                         */
/***************************************************************************/
int CFileInStream::Open(
#ifdef USE_WINAPI
  const TCHAR *pszFileName,
  bool bShareForWrite
#else
  const TCHAR *pszFileName
#endif  // USE_WINAPI
  )
{
  // Закрытие
  Close();

#ifdef USE_WINAPI
  DWORD dwShareMode = FILE_SHARE_READ;
  if (bShareForWrite) dwShareMode |= FILE_SHARE_WRITE;
  hFile = ::CreateFile(pszFileName, GENERIC_READ, dwShareMode, NULL,
                       OPEN_EXISTING, 0, NULL);
  return (hFile != INVALID_HANDLE_VALUE) ? 0 : (int)::GetLastError();
#else
  if (!_tfopen_s(&file, pszFileName, _T("rb")))
    return 0;
  return errno;
#endif  // USE_WINAPI
}
/***************************************************************************/
/* Close - Закрытие                                                        */
/***************************************************************************/
void CFileInStream::Close()
{
#ifdef USE_WINAPI
  if (hFile != INVALID_HANDLE_VALUE)
  {
    ::CloseHandle(hFile);
    hFile = INVALID_HANDLE_VALUE;
  }
#else
  if (file)
  {
    fclose(file);
    file = NULL;
  }
#endif  // USE_WINAPI
}
/***************************************************************************/
/* Read - Чтение данных                                                    */
/***************************************************************************/
int CFileInStream::Read(
  void *buf,
  size_t *size
  )
{
  if (*size == 0)
    return 0;

  size_t nBytesToRead = *size;

#ifdef USE_WINAPI
  DWORD dwBytesRead = 0;
  BOOL bRes = ::ReadFile(hFile, buf, (DWORD)nBytesToRead, &dwBytesRead,
                         NULL);
  *size = dwBytesRead;
  if (!bRes)
    return (int)::GetLastError();
#else
  *size = fread(buf, 1, nBytesToRead, file);
  if (*size != nBytesToRead)
    return ferror(file);
#endif  // USE_WINAPI
  return 0;
}
/***************************************************************************/
/* Seek - Установка указателя                                              */
/***************************************************************************/
int CFileInStream::Seek(
  unsigned __int64 *offset,
  StreamSeek origin
  )
{
#ifdef USE_WINAPI
  DWORD dwMoveMethod;
  DWORD dwOffsetLow = (DWORD)*offset;
  LONG lOffsetHigh = (LONG)(*offset >> 16 >> 16);
  switch (origin)
  {
    case STM_SEEK_SET: dwMoveMethod = FILE_BEGIN; break;
    case STM_SEEK_CUR: dwMoveMethod = FILE_CURRENT; break;
    case STM_SEEK_END: dwMoveMethod = FILE_END; break;
    default: return ERROR_INVALID_PARAMETER;
  }
  dwOffsetLow = ::SetFilePointer(hFile, dwOffsetLow, &lOffsetHigh,
                                 dwMoveMethod);
  if (dwOffsetLow == INVALID_SET_FILE_POINTER)
  {
    int res = ::GetLastError();
    if (res != NO_ERROR)
      return res;
  }
  *offset = ((unsigned __int64)lOffsetHigh << 32) | dwOffsetLow;
  return 0;
#else
  int moveMethod;
  int res;
  switch (origin)
  {
    case STM_SEEK_SET: moveMethod = SEEK_SET; break;
    case STM_SEEK_CUR: moveMethod = SEEK_CUR; break;
    case STM_SEEK_END: moveMethod = SEEK_END; break;
    default: return 1;
  }
  res = _fseeki64(file, *offset, moveMethod);
  *offset = (unsigned __int64)_ftelli64(file);
  return res;
#endif  // USE_WINAPI
}
//---------------------------------------------------------------------------
/***************************************************************************/
/* ~CFileInStream - Деструктор класса                                      */
/***************************************************************************/
CFileSeqOutStream::~CFileSeqOutStream()
{
  // Закрытие
  Close();
}
/***************************************************************************/
/* Open - Открытие                                                         */
/***************************************************************************/
int CFileSeqOutStream::Open(
  const TCHAR *pszFileName
  )
{
  // Закрытие
  Close();

#ifdef USE_WINAPI
  hFile = ::CreateFile(pszFileName, GENERIC_WRITE, FILE_SHARE_READ, NULL,
                       CREATE_ALWAYS, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
  return (hFile != INVALID_HANDLE_VALUE) ? 0 : (int)::GetLastError();
#else
  if (!_tfopen_s(&file, pszFileName, _T("wb+")))
    return 0;
  return errno;
#endif  // USE_WINAPI
}
/***************************************************************************/
/* Close - Закрытие                                                        */
/***************************************************************************/
void CFileSeqOutStream::Close()
{
#ifdef USE_WINAPI
  if (hFile != INVALID_HANDLE_VALUE)
  {
    ::CloseHandle(hFile);
    hFile = INVALID_HANDLE_VALUE;
  }
#else
  if (file)
  {
    fclose(file);
    file = NULL;
  }
#endif  // USE_WINAPI
}
/***************************************************************************/
/* Write - Запись данных                                                   */
/***************************************************************************/
size_t CFileSeqOutStream::Write(
  const void *data,
  size_t size
  )
{
  if (size == 0)
    return 0;

#ifdef USE_WINAPI
  DWORD dwBytesWritten = 0;
  BOOL bRes = ::WriteFile(hFile, data, (DWORD)size, &dwBytesWritten, NULL);
  return (size_t)dwBytesWritten;
#else
  return fwrite(data, 1, size, file);
#endif  // USE_WINAPI
}
//---------------------------------------------------------------------------
