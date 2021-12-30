//---------------------------------------------------------------------------
#ifndef __FILESTREAMS_H__
#define __FILESTREAMS_H__
//---------------------------------------------------------------------------
#ifdef USE_WINAPI
#define WIN32_LEAN_AND_MEAN  // Exclude rarely-used stuff from Windows headers
#include <windows.h>
#else
#include <stdio.h>
#endif  // USE_WINAPI

#include <stddef.h>
#include <tchar.h>
#include "Streams.h"
//---------------------------------------------------------------------------
// Класс файлового потока для ввода данных
class CFileInStream : public CInStream
{
public:
#ifdef USE_WINAPI
  HANDLE  hFile;
#else
  FILE *file;
#endif  // USE_WINAPI

  // Открытие
  int Open(
#ifdef USE_WINAPI
    const TCHAR *pszFileName,
    bool bShareForWrite = false
#else
    const TCHAR *pszFileName
#endif  // USE_WINAPI
    );
  // Закрытие
  void Close();
  // Чтение данных
  // (input(*size) != 0 && output(*size) == 0) -> EOF
  // (output(*size) < input(*size)) - разрешено
  virtual int Read(
    void *buf,
    size_t *size
    );
  // Установка указателя
  virtual int Seek(
    unsigned __int64 *offset,
    StreamSeek origin
    );

  // Конструктор класса
  CFileInStream() :
    CInStream(),
#ifdef USE_WINAPI
    hFile(INVALID_HANDLE_VALUE)
#else
    file(NULL)
#endif  // USE_WINAPI
    { }
  // Деструктор класса
  virtual ~CFileInStream();

private:
  // Копирование и присвоение не разрешены
  CFileInStream(const CFileInStream&);
  void operator=(const CFileInStream&);
};

// Класс файлового потока для последовательного вывода данных
class CFileSeqOutStream : public CSeqOutStream
{
public:
#ifdef USE_WINAPI
  HANDLE  hFile;
#else
  FILE *file;
#endif  // USE_WINAPI

  // Открытие
  int Open(
    const TCHAR *pszFileName
    );
  // Закрытие
  void Close();
  // Запись данных
  // result - количество записанных байтов
  // (result < size) -> ошибка
  virtual size_t Write(
    const void *data,
    size_t size
    );

  // Конструктор класса
  CFileSeqOutStream() :
    CSeqOutStream(),
#ifdef USE_WINAPI
    hFile(INVALID_HANDLE_VALUE)
#else
    file(NULL)
#endif  // USE_WINAPI
    { }
  // Деструктор класса
  virtual ~CFileSeqOutStream();

private:
  // Копирование и присвоение не разрешены
  CFileSeqOutStream(const CFileSeqOutStream&);
  void operator=(const CFileSeqOutStream&);
};
//---------------------------------------------------------------------------
#endif  // __FILESTREAMS_H__
