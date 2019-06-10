(ns perfect.poi.full
  (:require
   [perfect.poi.generics :as generics :reload true]
   [net.danielcompton.defn-spec-alpha :as ds]
   [clojure.spec.alpha :as s])
  (:import
   (org.apache.poi.xssf.usermodel XSSFWorkbook)
   (org.apache.poi.ss.usermodel Cell
                                Row
                                Sheet
                                Workbook)))

(ds/defn sheet
  [wb
   sheetn :- ::generics/sheet-identity]
  (s/assert ::generics/sheet-identity sheetn)
  (cond
    (pos-int? sheet) (.getSheetAt wb (dec sheetn))
    (string? sheet)  (.getSheet wb sheetn)))

(defprotocol POIMapper
  (from-java [obj] [obj opt]))

(extend-protocol POIMapper
  Workbook
  (from-java
    ([^Workbook wb]
     {:sheet-number (.getNumberOfSheets wb)
      :sheets       (into [] (map from-java) (sheets wb))})
    ([^Workbook wb sheetn]
     {:sheet-number (.getNumberOfSheets wb)
      :sheets       (from-java (.getSheetAt wb sheetn))}))

  Sheet
  (from-java [^Sheet sheet]
    {:first-row  (.getFirstRowNum sheet)
     :last-row   (.getLastRowNum sheet)
     :sheet-name (.getSheetName sheet)
     :first-col  (.getLeftCol sheet)
     :rows       (into []
                       (comp
                         (map from-java)
                         (filter #(not-empty (:cells %))))
                       (rows sheet))})

  Row
  (from-java [^Row row]
    {:n-cells    (.getPhysicalNumberOfCells row)
     :row-number (.getRowNum row)
     :first-col  (.getFirstCellNum row)
     :last-col   (.getLastCellNum row)
     :cells      (into []
                       (comp
                         (map from-java)
                         (keep identity))
                       (cells row))})

  Cell
  (from-java [^Cell cell]
    (let [cv (cell-value cell)]
      (when cv
        {:type    (str (.getCellType cell))
         :col-idx (.getColumnIndex cell)
         :value   cv}))))

(defn read-workbook
  ([path]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-java (XSSFWorkbook. ef))))
  ([path opt]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-java (XSSFWorkbook. ef) opt))))
