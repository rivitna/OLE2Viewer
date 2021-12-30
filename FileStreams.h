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
// ����� ��������� ������ ��� ����� ������
class CFileInStream : public CInStream
{
public:
#ifdef USE_WINAPI
  HANDLE  hFile;
#else
  FILE *file;
#endif  // USE_WINAPI

  // ��������
  int Open(
#ifdef USE_WINAPI
    const TCHAR *pszFileName,
    bool bShareForWrite = false
#else
    const TCHAR *pszFileName
#endif  // USE_WINAPI
    );
  // ��������
  void Close();
  // ������ ������
  // (input(*size) != 0 && output(*size) == 0) -> EOF
  // (output(*size) < input(*size)) - ���������
  virtual int Read(
    void *buf,
    size_t *size
    );
  // ��������� ���������
  virtual int Seek(
    unsigned __int64 *offset,
    StreamSeek origin
    );

  // ����������� ������
  CFileInStream() :
    CInStream(),
#ifdef USE_WINAPI
    hFile(INVALID_HANDLE_VALUE)
#else
    file(NULL)
#endif  // USE_WINAPI
    { }
  // ���������� ������
  virtual ~CFileInStream();

private:
  // ����������� � ���������� �� ���������
  CFileInStream(const CFileInStream&);
  void operator=(const CFileInStream&);
};

// ����� ��������� ������ ��� ����������������� ������ ������
class CFileSeqOutStream : public CSeqOutStream
{
public:
#ifdef USE_WINAPI
  HANDLE  hFile;
#else
  FILE *file;
#endif  // USE_WINAPI

  // ��������
  int Open(
    const TCHAR *pszFileName
    );
  // ��������
  void Close();
  // ������ ������
  // result - ���������� ���������� ������
  // (result < size) -> ������
  virtual size_t Write(
    const void *data,
    size_t size
    );

  // ����������� ������
  CFileSeqOutStream() :
    CSeqOutStream(),
#ifdef USE_WINAPI
    hFile(INVALID_HANDLE_VALUE)
#else
    file(NULL)
#endif  // USE_WINAPI
    { }
  // ���������� ������
  virtual ~CFileSeqOutStream();

private:
  // ����������� � ���������� �� ���������
  CFileSeqOutStream(const CFileSeqOutStream&);
  void operator=(const CFileSeqOutStream&);
};
//---------------------------------------------------------------------------
#endif  // __FILESTREAMS_H__
