//---------------------------------------------------------------------------
#ifndef __STREAMS_H__
#define __STREAMS_H__
//---------------------------------------------------------------------------
#include <stddef.h>
//---------------------------------------------------------------------------
// ����� ������ ��� ����������������� ����� ������
class CSeqInStream
{
public:
  // ������ ������
  // (input(*size) != 0 && output(*size) == 0) -> EOF
  // (output(*size) < input(*size)) - ���������
  virtual int Read(
    void *buf,
    size_t *size
    ) = 0;

  virtual ~CSeqInStream() { }
};

typedef enum
{
  STM_SEEK_SET = 0,
  STM_SEEK_CUR = 1,
  STM_SEEK_END = 2
} StreamSeek;

// ����� ������ ��� ����� ������
class CInStream
{
public:
  // ������ ������
  // (input(*size) != 0 && output(*size) == 0) -> EOF
  // (output(*size) < input(*size)) - ���������
  virtual int Read(
    void *buf,
    size_t *size
    ) = 0;
  // ��������� ���������
  virtual int Seek(
    unsigned __int64 *offset,
    StreamSeek origin
    ) = 0;

  virtual ~CInStream() { }
};

// ����� ������ ��� ����������������� ������ ������
class CSeqOutStream
{
public:
  // ������ ������
  // result - ���������� ���������� ������
  // (result < size) -> ������
  virtual size_t Write(
    const void *data,
    size_t size
    ) = 0;

  virtual ~CSeqOutStream() { }
};
//---------------------------------------------------------------------------
#endif  // __STREAMS_H__
