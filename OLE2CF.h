//---------------------------------------------------------------------------
#ifndef __OLE2CF_H__
#define __OLE2CF_H__
//---------------------------------------------------------------------------
#include "OLE2Type.h"
//---------------------------------------------------------------------------
// Сигнатура составного файла OLE2
#define OLE2_SIGNATURE  0xE011CFD0
#define OLE2_VERSION    0xE11AB1A1
//---------------------------------------------------------------------------
// Маркеры секторов в FAT
// Регулярный номер сектора REGSECT = 0x00000000 - 0xFFFFFFF9
#define MAXREGSECT  0xFFFFFFFA       // Максимальный регулярный номер сектора
#define DIFSECT     0xFFFFFFFC       // Сектор DIFAT
#define FATSECT     0xFFFFFFFD       // Сектор FAT
#define ENDOFCHAIN  0xFFFFFFFE       // Конец цепочки секторов
#define FREESECT    0xFFFFFFFF       // Свободный сектор
//---------------------------------------------------------------------------
#pragma pack(push, 1)
//---------------------------------------------------------------------------
/***************************************************************************/
/* Заголовок составного файла OLE2                                         */
/***************************************************************************/

// Количество записей DIFAT в заголовке
#define NUMBER_OF_DIFAT_ENTRIES  109

typedef struct _COMPOUND_FILE_HEADER
{
  uint32_t  dwSign;                  // 0000: Сигнатура
  uint32_t  dwVersion;               // 0004: Версия
  uint8_t   bCLSID[16];              // 0008: CLSID
                                     //       (=CLSID_NULL)
  uint16_t  wMinorVersion;           // 0018: Младший номер версии
                                     //       (=0x003E)
  uint16_t  wMajorVersion;           // 001A: Старший номер версии
                                     //       (=0x0003 или =0x0004)
  uint16_t  wByteOrder;              // 001C: Порядок байтов
                                     //       (=0xFFFE - little-endian)
  uint16_t  wSectorShift;            // 001E: LOG2(размер сектора)
                                     //       (=0x0009 - wMajorVersion=3,
                                     //        =0x000C - wMajorVersion=4)
  uint16_t  wMiniSectorShift;        // 0020: LOG2(размер сектора mini FAT)
                                     //       (=0x0006)
  uint8_t   bReserved[6];            // 0022: Зарезервировано
  uint32_t  dwNumDirSectors;         // 0028: Количество секторов каталога
                                     //       в файле
                                     //       (=0 - wMajorVersion=3)
  uint32_t  dwNumFATSectors;         // 002C: Количество секторов FAT в файле
  uint32_t  dwDirStartSector;        // 0030: Номер начального сектора
                                     //       каталога
  uint32_t  dwTransactionSignNum;    // 0034: Номер сигнатуры транзакции
                                     //       файла
  uint32_t  dwMiniStreamCutoffSize;  // 0038: Максимальный размер потока,
                                     //       размещаемого в mini FAT
                                     //       (=0x00001000)
  uint32_t  dwMiniFATStartSector;    // 003C: Номер начального сектора
                                     //       mini FAT
  uint32_t  dwNumMiniFATSectors;     // 0040: Количество секторов mini FAT
                                     //       в файле
  uint32_t  dwDIFATStartSector;      // 0044: Номер начального сектора DIFAT
  uint32_t  dwNumDIFATSectors;       // 0048: Количество секторов DIFAT
                                     //       в файле
  uint32_t  dwDIFAT[NUMBER_OF_DIFAT_ENTRIES];  // 004C: DIFAT
} COMPOUND_FILE_HEADER, *PCOMPOUND_FILE_HEADER;

#define COMPOUND_FILE_HEADER_SIZE  sizeof(COMPOUND_FILE_HEADER)

/***************************************************************************/
/* Запись каталога                                                         */
/***************************************************************************/

// Максимальный размер имени записи каталога
#define MAX_DIR_ENTRY_NAME_SIZE  64

typedef struct _DIRECTORY_ENTRY
{
  uint8_t   bName[MAX_DIR_ENTRY_NAME_SIZE];  // 0000: Имя (UTF-16)
  uint16_t  wSizeOfName;             // 0040: Размер имени в байтах
  uint8_t   bObjectType;             // 0042: Тип объекта
  uint8_t   bColorFlag;              // 0043: Флаг цвета (красный/черный)
  uint32_t  dwLeftSiblingID;         // 0044: Идентификатор одноуровневого
                                     //       объекта слева
  uint32_t  dwRightSiblingID;        // 0048: Идентификатор одноуровневого
                                     //       объекта справа
  uint32_t  dwChildID;               // 004C: Идентификатор дочернего объекта
  uint8_t   bCLSID[16];              // 0050: CLSID
  uint32_t  dwStateBits;             // 0060: Флаги состояния
  union                              // 0064: Дата создания (FILETIME)
  {
    uint8_t  bCreationTime[8];
    struct
    {
      uint32_t  dwLowCreationTime;
      uint32_t  dwHighCreationTime;
    };
  };
  union                              // 006С: Дата изменения (FILETIME)
  {
    uint8_t  bModifiedTime[8];
    struct
    {
      uint32_t  dwLowModifiedTime;
      uint32_t  dwHighModifiedTime;
    };
  };
  uint32_t  dwStartSector;           // 0074: Номер начального сектора
  union                              // 0078: Размер потока
  {                                  //       (<= 0x80000000 - версия
    uint64_t  qwStreamSize;          //                        файла 3)
    struct
    {
      uint32_t  dwLowStreamSize;
      uint32_t  dwHighStreamSize;
    };
  };
} DIRECTORY_ENTRY, *PDIRECTORY_ENTRY;

#define DIRECTORY_ENTRY_SIZE  sizeof(DIRECTORY_ENTRY)

// Типы объектов
#define OBJ_TYPE_UNKNOWN  0x00       // Неизвестный или не используется
#define OBJ_TYPE_STORAGE  0x01       // Каталог
#define OBJ_TYPE_STREAM   0x02       // Поток
#define OBJ_TYPE_ROOT     0x05       // Корневой каталог

// Флаги цвета объектов
#define OBJ_COLOR_RED     0x00       // Красный
#define OBJ_COLOR_BLACK   0x01       // Черный

// Идентификаторы объектов
// 0x00000000 <= object ID <= 0xFFFFFFF9
#define MAXREGSID  0xFFFFFFFA        // Максимальный регулярный идентификатор
                                     // потока
#define NOSTREAM   0xFFFFFFFF        // Объект отсутствует

// Максимальный размер потока для версии составного файла 3
#define V3_MAX_STREAM_SIZE  0x80000000
//---------------------------------------------------------------------------
#pragma pack(pop)
//---------------------------------------------------------------------------
#endif  // __OLE2CF_H__
