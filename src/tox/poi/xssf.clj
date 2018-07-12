(ns tox.poi.xssf
  (:import (java.io FileOutputStream FileInputStream InputStream OutputStream)(junit.framework.Assert)
           (org.apache.poi.hssf.usermodel HSSFWorkbook)
           (org.apache.poi.xssf.usermodel XSSFWorkbook)
           (org.apache.poi.ss.usermodel 
            Workbook Sheet Cell Row
            FormulaError
            WorkbookFactory DateUtil
            IndexedColors CellStyle Font
            CellValue Drawing CreationHelper)
           (org.apache.poi.ss.util CellReference)
           (org.apache.poi.xssf.streaming SXSSFWorkbook)))
(let [wb (HSSFWorkbook.)
      createHelper (.getCreationHelper wb)
      s (.createSheet wb)
      cs (.createCellStyle wb)
      cs2 (.createCellStyle wb)
      df (.createDataFormat wb)
      f (.createFont wb)
      f2 (.createFont wb)]
  (doto f 
    (.setFontHeightInPoints (short 12))
    (.setColor (.getIndex (IndexedColors/valueOf "RED")))
    (.setBold true)
   )
  (doto cs
    (.setFont f)
    (.setDataFormat (.getFormat df "#,##0.0")))
  (doseq [rownum (range 30)]
    (let [r (.createRow s rownum)]
      (doseq [cellnum (range 10)]
        (doto (.createCell r cellnum)
          (.setCellValue (+ (double rownum) (/ cellnum 10)))))))
  (with-open [file-out (FileOutputStream. "abc.xls")]
    (.write wb file-out)))
