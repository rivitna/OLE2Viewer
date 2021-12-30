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
// Проверка, является ли указанный путь относительным
bool IsRelativePath(
  const TCHAR *pszPath
  );
// Проверка существования файла
bool IsFileExists(
  const TCHAR *pszFilePath
  );
// Проверка существования каталога
bool IsDirectoryExists(
  const TCHAR *pszDirPath
  );
// Создание иерархии каталогов
bool ForceDirectories(
  const TCHAR *pszDirPath
  );
// Получение имени файла
TCHAR *GetFileName(
  const TCHAR *pszFilePath
  );
// Извлечение пути к файлу
size_t ExtractFilePath(
  const TCHAR *pszFileName,
  TCHAR *pBuffer,
  size_t nSize
  );
// Получение расширения в имени файла
TCHAR *GetFileExt(
  const TCHAR *pszFilePath
  );
// Добавление расширения к имени файла
size_t AddFileExt(
  const TCHAR *pszFileName,
  const TCHAR *pszExt,
  TCHAR *pBuffer,
  size_t nSize
  );
// Изменение расширения в имени файла
size_t ChangeFileExt(
  const TCHAR *pszFileName,
  const TCHAR *pszExt,
  TCHAR *pBuffer,
  size_t nSize
  );
// Объединение путей
size_t CombinePath(
  const TCHAR *pszDir,
  const TCHAR *pszFile,
  TCHAR *pBuffer,
  size_t nSize
  );
// Объединение путей
// (Возвращенный указатель необходимо освободить при помощи функции free)
TCHAR *AllocAndCombinePath(
  const TCHAR *pszDir,
  const TCHAR *pszFile
  );
// Изменение имени файла
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
