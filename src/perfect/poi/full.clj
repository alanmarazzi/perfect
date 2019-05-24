(ns perfect.poi.full
  (:require
   [perfect.poi.generics :refer [cell-value
                              sheets
                              rows
                              cells
                              values]]
   [com.climate.claypoole :as cp])
  (:import
   (org.apache.poi.ss.usermodel Cell
                                Row
                                Sheet
                                Workbook)))

(defprotocol POIMapper
  (from-java [obj]))

(extend-protocol POIMapper
  Workbook
  (from-java [^Workbook wb]
    {:sheet-number (.getNumberOfSheets wb)
     :sheets       ;(map from-java (sheets wb))
     (into [] (map from-java) (sheets wb))
     :level        :workbook})

  Sheet
  (from-java [^Sheet sheet]
    {:first-row  (.getFirstRowNum sheet)
     :last-row   (.getLastRowNum sheet)
     :sheet-name (.getSheetName sheet)
     :first-col  (.getLeftCol sheet)
     :rows       ;(cp/upmap 4 from-java (rows sheet))
     (into [] (map from-java) (rows sheet))
     :level      :sheet})

  Row
  (from-java [^Row row]
    {:n-cells    (.getPhysicalNumberOfCells row)
     :row-number (.getRowNum row)
     :first-col  (.getFirstCellNum row)
     :last-col   (.getLastCellNum row)
     :cells      ;(pmap from-java (cells row))
                                        (into [] (map from-java) (cells row))
     :level      :row})

  Cell
  (from-java [^Cell cell]
    {:type    (str (.getCellType cell))
     :row-idx (.getRowIndex cell)
     :col-idx (.getColumnIndex cell)
     :value   (cell-value cell)
     :level   :cell}))
