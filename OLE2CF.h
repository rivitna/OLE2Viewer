//---------------------------------------------------------------------------
#ifndef __OLE2CF_H__
#define __OLE2CF_H__
//---------------------------------------------------------------------------
#include "OLE2Type.h"
//---------------------------------------------------------------------------
// ��������� ���������� ����� OLE2
#define OLE2_SIGNATURE  0xE011CFD0
#define OLE2_VERSION    0xE11AB1A1
//---------------------------------------------------------------------------
// ������� �������� � FAT
// ���������� ����� ������� REGSECT = 0x00000000 - 0xFFFFFFF9
#define MAXREGSECT  0xFFFFFFFA       // ������������ ���������� ����� �������
#define DIFSECT     0xFFFFFFFC       // ������ DIFAT
#define FATSECT     0xFFFFFFFD       // ������ FAT
#define ENDOFCHAIN  0xFFFFFFFE       // ����� ������� ��������
#define FREESECT    0xFFFFFFFF       // ��������� ������
//---------------------------------------------------------------------------
#pragma pack(push, 1)
//---------------------------------------------------------------------------
/***************************************************************************/
/* ��������� ���������� ����� OLE2                                         */
/***************************************************************************/

// ���������� ������� DIFAT � ���������
#define NUMBER_OF_DIFAT_ENTRIES  109

typedef struct _COMPOUND_FILE_HEADER
{
  uint32_t  dwSign;                  // 0000: ���������
  uint32_t  dwVersion;               // 0004: ������
  uint8_t   bCLSID[16];              // 0008: CLSID
                                     //       (=CLSID_NULL)
  uint16_t  wMinorVersion;           // 0018: ������� ����� ������
                                     //       (=0x003E)
  uint16_t  wMajorVersion;           // 001A: ������� ����� ������
                                     //       (=0x0003 ��� =0x0004)
  uint16_t  wByteOrder;              // 001C: ������� ������
                                     //       (=0xFFFE - little-endian)
  uint16_t  wSectorShift;            // 001E: LOG2(������ �������)
                                     //       (=0x0009 - wMajorVersion=3,
                                     //        =0x000C - wMajorVersion=4)
  uint16_t  wMiniSectorShift;        // 0020: LOG2(������ ������� mini FAT)
                                     //       (=0x0006)
  uint8_t   bReserved[6];            // 0022: ���������������
  uint32_t  dwNumDirSectors;         // 0028: ���������� �������� ��������
                                     //       � �����
                                     //       (=0 - wMajorVersion=3)
  uint32_t  dwNumFATSectors;         // 002C: ���������� �������� FAT � �����
  uint32_t  dwDirStartSector;        // 0030: ����� ���������� �������
                                     //       ��������
  uint32_t  dwTransactionSignNum;    // 0034: ����� ��������� ����������
                                     //       �����
  uint32_t  dwMiniStreamCutoffSize;  // 0038: ������������ ������ ������,
                                     //       ������������ � mini FAT
                                     //       (=0x00001000)
  uint32_t  dwMiniFATStartSector;    // 003C: ����� ���������� �������
                                     //       mini FAT
  uint32_t  dwNumMiniFATSectors;     // 0040: ���������� �������� mini FAT
                                     //       � �����
  uint32_t  dwDIFATStartSector;      // 0044: ����� ���������� ������� DIFAT
  uint32_t  dwNumDIFATSectors;       // 0048: ���������� �������� DIFAT
                                     //       � �����
  uint32_t  dwDIFAT[NUMBER_OF_DIFAT_ENTRIES];  // 004C: DIFAT
} COMPOUND_FILE_HEADER, *PCOMPOUND_FILE_HEADER;

#define COMPOUND_FILE_HEADER_SIZE  sizeof(COMPOUND_FILE_HEADER)

/***************************************************************************/
/* ������ ��������                                                         */
/***************************************************************************/

// ������������ ������ ����� ������ ��������
#define MAX_DIR_ENTRY_NAME_SIZE  64

typedef struct _DIRECTORY_ENTRY
{
  uint8_t   bName[MAX_DIR_ENTRY_NAME_SIZE];  // 0000: ��� (UTF-16)
  uint16_t  wSizeOfName;             // 0040: ������ ����� � ������
  uint8_t   bObjectType;             // 0042: ��� �������
  uint8_t   bColorFlag;              // 0043: ���� ����� (�������/������)
  uint32_t  dwLeftSiblingID;         // 0044: ������������� ��������������
                                     //       ������� �����
  uint32_t  dwRightSiblingID;        // 0048: ������������� ��������������
                                     //       ������� ������
  uint32_t  dwChildID;               // 004C: ������������� ��������� �������
  uint8_t   bCLSID[16];              // 0050: CLSID
  uint32_t  dwStateBits;             // 0060: ����� ���������
  union                              // 0064: ���� �������� (FILETIME)
  {
    uint8_t  bCreationTime[8];
    struct
    {
      uint32_t  dwLowCreationTime;
      uint32_t  dwHighCreationTime;
    };
  };
  union                              // 006�: ���� ��������� (FILETIME)
  {
    uint8_t  bModifiedTime[8];
    struct
    {
      uint32_t  dwLowModifiedTime;
      uint32_t  dwHighModifiedTime;
    };
  };
  uint32_t  dwStartSector;           // 0074: ����� ���������� �������
  union                              // 0078: ������ ������
  {                                  //       (<= 0x80000000 - ������
    uint64_t  qwStreamSize;          //                        ����� 3)
    struct
    {
      uint32_t  dwLowStreamSize;
      uint32_t  dwHighStreamSize;
    };
  };
} DIRECTORY_ENTRY, *PDIRECTORY_ENTRY;

#define DIRECTORY_ENTRY_SIZE  sizeof(DIRECTORY_ENTRY)

// ���� ��������
#define OBJ_TYPE_UNKNOWN  0x00       // ����������� ��� �� ������������
#define OBJ_TYPE_STORAGE  0x01       // �������
#define OBJ_TYPE_STREAM   0x02       // �����
#define OBJ_TYPE_ROOT     0x05       // �������� �������

// ����� ����� ��������
#define OBJ_COLOR_RED     0x00       // �������
#define OBJ_COLOR_BLACK   0x01       // ������

// �������������� ��������
// 0x00000000 <= object ID <= 0xFFFFFFF9
#define MAXREGSID  0xFFFFFFFA        // ������������ ���������� �������������
                                     // ������
#define NOSTREAM   0xFFFFFFFF        // ������ �����������

// ������������ ������ ������ ��� ������ ���������� ����� 3
#define V3_MAX_STREAM_SIZE  0x80000000
//---------------------------------------------------------------------------
#pragma pack(pop)
//---------------------------------------------------------------------------
#endif  // __OLE2CF_H__
