//---------------------------------------------------------------------------
#ifndef __STREAMS_H__
#define __STREAMS_H__
//---------------------------------------------------------------------------
#include <stddef.h>
//---------------------------------------------------------------------------
// Класс потока для последовательного ввода данных
class CSeqInStream
{
public:
  // Чтение данных
  // (input(*size) != 0 && output(*size) == 0) -> EOF
  // (output(*size) < input(*size)) - разрешено
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

// Класс потока для ввода данных
class CInStream
{
public:
  // Чтение данных
  // (input(*size) != 0 && output(*size) == 0) -> EOF
  // (output(*size) < input(*size)) - разрешено
  virtual int Read(
    void *buf,
    size_t *size
    ) = 0;
  // Установка указателя
  virtual int Seek(
    unsigned __int64 *offset,
    StreamSeek origin
    ) = 0;

  virtual ~CInStream() { }
};

// Класс потока для последовательного вывода данных
class CSeqOutStream
{
public:
  // Запись данных
  // result - количество записанных байтов
  // (result < size) -> ошибка
  virtual size_t Write(
    const void *data,
    size_t size
    ) = 0;

  virtual ~CSeqOutStream() { }
};
//---------------------------------------------------------------------------
#endif  // __STREAMS_H__
