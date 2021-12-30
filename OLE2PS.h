//---------------------------------------------------------------------------
#ifndef __OLE2PS_H__
#define __OLE2PS_H__
//---------------------------------------------------------------------------
#include "OLE2Type.h"
//---------------------------------------------------------------------------
#pragma pack(push, 1)
#pragma warning(disable: 4200)
//---------------------------------------------------------------------------
/***************************************************************************/
/* Запись набора свойств                                                   */
/***************************************************************************/
typedef struct _PROPERTYSET_ENTRY
{
  uint8_t            bFMTID[16];          // 0000: Идентификатор формата
                                          //       пакета набора свойств
                                          //       (GUID)
                                          //       FMTID_SummaryInformation
                                          //       ({F29F85E0-4FF9-1068-AB91-
                                          //         08002B27B3D9})
                                          //       FMTID_DocSummaryInformation
                                          //       ({D5CDD502-2E9C-101B-9397-
                                          //         08002B2CF9AE})
                                          //       FMTID_UserDefinedProperties
                                          //       ({D5CDD505-2E9C-101B-9397-
                                          //         08002B2CF9AE}
  uint32_t           dwOffset;            // 0010: Смещение набора свойств
} PROPERTYSET_ENTRY, *PPROPERTYSET_ENTRY;

/***************************************************************************/
/* Заголовок потока наборов свойств                                        */
/***************************************************************************/
typedef struct _PROPERTYSET_STREAM_HEADER
{
  uint16_t           wByteOrder;          // 0000: Порядок байтов
                                          //       (=0xFFFE - little-endian)
  uint16_t           wVersion;            // 0002: Версия (=0 или =1)
  uint16_t           wOSVersion;          // 0004: Версия ОС
  uint16_t           wOSType;             // 0006: Тип ОС
                                          //       (=0 - Win16
                                          //        =1 - Mac
                                          //        =2 - Win32)
  uint8_t            bClassID[16];        // 0008: Class ID потока
                                          //       (=CLSID_NULL)
  uint32_t           dwNumPropSets;       // 0018: Количество наборов свойств
  PROPERTYSET_ENTRY  propSetEntries[0];   // 001C: Записи наборов свойств
} PROPERTYSET_STREAM_HEADER, *PPROPERTYSET_STREAM_HEADER;

/***************************************************************************/
/* Вариантное значение свойства                                            */
/***************************************************************************/
typedef struct _PROPERTY_VARIANT
{
  uint16_t           vt;                  // Тип
  uint16_t           wPadding;            // Выравнивание до 4 байтов
  union
  {
    uint16_t         wVal;                // Значение PVT_I2
    uint32_t         dwVal;               // Значение PVT_I4
    uint16_t         bVal;                // Значение PVT_BOOL
    uint32_t         cb;                  // Число байтов данных
                                          // (PVT_LPSTR, PVT_BLOB, PVT_CF)
    uint32_t         nItemCount;          // Количество элементов вектора
                                          // (PVT_VARIANT_VECTOR,
                                          // PVT_LPSTR_VECTOR)
    struct                                // Значение PVT_FILETIME
    {
      uint32_t       dwLowDateTime;
      uint32_t       dwHighDateTime;
    } ftVal;
    uint64_t         qwVal;               // 64-битное значение
  };
} PROPERTY_VARIANT, *PPROPERTY_VARIANT;

// Минимальный размер значения свойства
#define MIN_PROP_VAR_SIZE  (2 * sizeof(uint32_t))

// Смещение данных значения свойства
#define PROP_VAR_DATA_OFFSET  (2 * sizeof(uint32_t))

// Данные значения свойства
#define PROP_VAR_DATA(pPropVar)  \
  ((uint8_t *)(pPropVar) + PROP_VAR_DATA_OFFSET)
// Строчное значение свойства
#define PROP_VAR_LPSTR(pPropVar)  ((char *)PROP_VAR_DATA(pPropVar))

// Типы свойств
#define PVT_I2                0x0002      // = VT_I2
#define PVT_I4                0x0003      // = VT_I4
#define PVT_BOOL              0x000B      // = VT_BOOL
#define PVT_VARIANT           0x000C      // = VT_VARIANT
#define PVT_UI4               0x0013      // = VT_UI4
#define PVT_LPSTR             0x001E      // = VT_LPSTR
#define PVT_FILETIME          0x0040      // = VT_FILETIME
#define PVT_BLOB              0x0041      // = VT_BLOB
#define PVT_CF                0x0047      // = VT_CF
#define PVT_VARIANT_VECTOR    0x100C      // = VT_VECTOR | VT_VARIANT
#define PVT_LPSTR_VECTOR      0x101E      // = VT_VECTOR | VT_LPSTR

// Идентификаторы свойств
#define PID_DICTIONARY        0x00000000  // Dictionary property.
#define PID_CODEPAGE          0x00000001  // (PVT_I2) Code page property.
#define PID_LOCALEID          0x80000000  // (PVT_UI4) Locale ID property.
#define PID_BEHAVIOR          0x80000003  // (PVT_UI4) Behavior Property.
                                          // If the Behavior property is
                                          // present, it MUST have one of the
                                          // following values:
                                          // 0x00000001: Property names are
                                          //             case-insensitive
                                          //             (default).
                                          // 0x00000002: Property names are
                                          //             case-sensitive.

// Идентификаторы свойств набора формата FMTID_SummaryInformation
#define PIDSI_TITLE           0x00000002  // (PVT_LPSTR) Title.
#define PIDSI_SUBJECT         0x00000003  // (PVT_LPSTR) Subject.
#define PIDSI_AUTHOR          0x00000004  // (PVT_LPSTR) Author.
#define PIDSI_KEYWORDS        0x00000005  // (PVT_LPSTR) Keywords.
#define PIDSI_COMMENTS        0x00000006  // (PVT_LPSTR) Comments.
#define PIDSI_TEMPLATE        0x00000007  // (PVT_LPSTR) Template.
#define PIDSI_LASTAUTHOR      0x00000008  // (PVT_LPSTR) Last Saved By.
#define PIDSI_REVNUMBER       0x00000009  // (PVT_LPSTR) Revision Number.
#define PIDSI_EDITTIME        0x0000000A  // (PVT_FILETIME) Edit Time.
#define PIDSI_LASTPRINTED     0x0000000B  // (PVT_FILETIME) Last Printed.
#define PIDSI_CREATE_DTM      0x0000000C  // (PVT_FILETIME) Date Created.
#define PIDSI_LASTSAVE_DTM    0x0000000D  // (PVT_FILETIME) Date Last Saved.
#define PIDSI_PAGECOUNT       0x0000000E  // (PVT_I4) Page Count.
#define PIDSI_WORDCOUNT       0x0000000F  // (PVT_I4) Word Count.
#define PIDSI_CHARCOUNT       0x00000010  // (PVT_I4) Character Count.
#define PIDSI_THUMBNAIL       0x00000011  // (PVT_CF) Application-specific
                                          // clipboard data containing a
                                          // thumbnail representing the
                                          // document's contents.
#define PIDSI_APPNAME         0x00000012  // (PVT_LPSTR) Application Name.
#define PIDSI_DOC_SECURITY    0x00000013  // (PVT_I4) Set of application-
                                          // suggested access control flags
                                          // with the following values:
                                          // 0x00000001: Password protected
                                          // 0x00000002: Read-only
                                          //             recommended
                                          // 0x00000004: Read-only enforced
                                          // 0x00000008: Locked for
                                          //             annotations

// Идентификаторы свойств набора формата FMTID_DocSummaryInformation
#define PIDDSI_CATEGORY       0x00000002  // (PVT_LPSTR) Category.
#define PIDDSI_PRESFORMAT     0x00000003  // (PVT_LPSTR) The presentation
                                          // format type of the document.
#define PIDDSI_BYTECOUNT      0x00000004  // (PVT_I4) Byte Count.
#define PIDDSI_LINECOUNT      0x00000005  // (PVT_I4) Line Count.
#define PIDDSI_PARACOUNT      0x00000006  // (PVT_I4) Paragragh Count.
#define PIDDSI_SLIDECOUNT     0x00000007  // (PVT_I4) Slides.
#define PIDDSI_NOTECOUNT      0x00000008  // (PVT_I4) Notes.
#define PIDDSI_HIDDENCOUNT    0x00000009  // (PVT_I4) Hidden Slides.
#define PIDDSI_MMCLIPCOUNT    0x0000000A  // (PVT_I4) Multimedia Clips.
#define PIDDSI_SCALE          0x0000000B  // (PVT_BOOL) The value of the
                                          // property MUST be FALSE.
#define PIDDSI_HEADINGPAIR    0x0000000C  // (PVT_VARIANT_VECTOR) Heading
                                          // string and a count of document
                                          // parts as found in the
                                          // PIDDSI_DOCPARTS property to
                                          // which this heading applies. The
                                          // total sum of document counts for
                                          // all headers in this property
                                          // MUST be equal to the number of
                                          // elements in the PIDDSI_DOCPARTS
                                          // property.
#define PIDDSI_DOCPARTS       0x0000000D  // (PVT_LPSTR_VECTOR) Each string
                                          // element of the vector specifies
                                          // a part of the document.The
                                          // elements of this vector are
                                          // ordered according to the header
                                          // they belong to as defined in the
                                          // PIDDSI_HEADINGPAIR property.
#define PIDDSI_MANAGER        0x0000000E  // (PVT_LPSTR) Manager.
#define PIDDSI_COMPANY        0x0000000F  // (PVT_LPSTR) Company.
#define PIDDSI_LINKSDIRTY     0x00000010  // (PVT_BOOL) The property value
                                          // specifies TRUE if any of the
                                          // values for the linked properties
                                          // in the User Defined Property Set
                                          // have changed outside of the
                                          // application, which would require
                                          // the application to update the
                                          // linked fields on document load.
#define PIDDSI_CCHWITHSPACES  0x00000011  // (PVT_I4) Estimate of the number
                                          // of characters in the document,
                                          // including whitespace.
#define PIDDSI_SHAREDDOC      0x00000013  // (PVT_BOOL) The property value
                                          // MUST be FALSE.
#define PIDDSI_LINKBASE       0x00000014  // MUST NOT be written.
#define PIDDSI_HLINKS         0x00000015  // MUST NOT be written.
#define PIDDSI_HYPERLINKSCHANGED  0x00000016  // (PVT_BOOL) The property
                                              // value specifies TRUE if the
                                              // _PID_HLINKS property in the
                                              // User Defined Property Set
                                              // has changed outside of the
                                              // application, which would
                                              // require the application to
                                              // update the hyperlink on
                                              // document load.
#define PIDDSI_VERSION        0x00000017  // (PVT_I4) Specifies the version
                                          // of the application that wrote
                                          // the property set storage.
#define PIDDSI_DIGSIG         0x00000018  // (PVT_BLOB) The data of the VBA
                                          // digital signature for the VBA
                                          // project embedded in the
                                          // document. MUST NOT exist if the
                                          // VBA project of the document does
                                          // not have a digital signature or
                                          // if the project is absent.
#define PIDDSI_CONTENTTYPE    0x0000001A  // (PVT_LPSTR) The content type of
                                          // the file.
#define PIDDSI_CONTENTSTATUS  0x0000001B  // (PVT_LPSTR) The document status.
#define PIDDSI_LANGUAGE       0x0000001C  // (PVT_LPSTR) SHOULD be absent.
#define PIDDSI_DOCVERSION     0x0000001D  // (PVT_LPSTR) SHOULD be absent.

/***************************************************************************/
/* Запись свойства                                                         */
/***************************************************************************/
typedef struct _PROPERTY_ENTRY
{
  uint32_t           dwID;                // 0000: Идентификатор свойства
  uint32_t           dwOffset;            // 0004: Смещение свойства от
                                          //       начала пакета набора
                                          //       свойств
} PROPERTY_ENTRY, *PPROPERTY_ENTRY;

/***************************************************************************/
/* Заголовок набора свойств                                                */
/***************************************************************************/
typedef struct _PROPERTYSET_HEADER
{
  uint32_t           dwSize;              // 0000: Размер пакета набора
                                          //       свойств
  uint32_t           dwNumProperties;     // 0004: Количество свойств
  PROPERTY_ENTRY     propEntries[0];      // 0008: Записи свойств
} PROPERTYSET_HEADER, *PPROPERTYSET_HEADER;

/***************************************************************************/
/* Запись словаря                                                          */
/***************************************************************************/
typedef struct _DICTIONARY_ENTRY
{
  uint32_t           dwPropID;            // 0000: Идентификатор свойства
  uint32_t           dwPropNameLength;    // 0004: Длина имени свойства
  uint8_t            bPropName;           // 0008: Имя свойства
                                          //       (если CodePage=0x04B0,
                                          //        используются 16-битные
                                          //        символы Unicode)
} DICTIONARY_ENTRY, *PDICTIONARY_ENTRY_ENTRY;
//---------------------------------------------------------------------------
#pragma pack(pop)
//---------------------------------------------------------------------------
#endif  // __OLE2PS_H__
