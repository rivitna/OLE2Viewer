//---------------------------------------------------------------------------
#ifndef __OLE2TYPE_H__
#define __OLE2TYPE_H__
//---------------------------------------------------------------------------
#include <stddef.h>
//---------------------------------------------------------------------------
#ifndef countof
#define countof(a)  (sizeof(a)/sizeof(a[0]))
#endif  // countof

#ifndef offsetof
#define offsetof(type, field)  ((size_t)&(((type *)0)->field))
#endif  // offsetof

#ifndef sizeof_through_field
#define sizeof_through_field(type, field)  \
  ((size_t)&(((type *)0)->field) + sizeof(((type *)0)->field))
#endif  // sizeof_through_field
//---------------------------------------------------------------------------
// ������������� ����
#ifdef _MSC_VER
#if _MSC_VER >= 1600
/* VS2010 */
#include <stdint.h>
#else
typedef char    int8_t;
typedef short   int16_t;
typedef long    int32_t;
typedef __int64 int64_t;

typedef unsigned char    uint8_t;
typedef unsigned short   uint16_t;
typedef unsigned long    uint32_t;
typedef unsigned __int64 uint64_t;
#endif  // _MSC_VER >= 1600
#endif  // _MSC_VER

// ���� ��������
#define O2_OK                 0  // �������
#define O2_ERROR_PARAM        1  // �������� ��������
#define O2_ERROR_UNSUPPORTED  2  // �� ��������������
#define O2_ERROR_MEM          3  // ������������ ������
#define O2_ERROR_SIGN         4  // �������� ���������
#define O2_ERROR_DATA         5  // �������� ������ ������
#define O2_ERROR_OPEN         6  // ������ ��������
#define O2_ERROR_READ         7  // ������ ������ ������
#define O2_ERROR_WRITE        8  // ������ ������ ������
#define O2_ERROR_INTERNAL     9  // ���������� ������

// ��� ���� ��������
typedef int O2Res;

// �������� ������
typedef struct _O2DataRange
{
  size_t  nOffset;
  size_t  nLength;
} O2DataRange, *PO2DataRange;
//---------------------------------------------------------------------------
#endif  // __OLE2TYPE_H__
