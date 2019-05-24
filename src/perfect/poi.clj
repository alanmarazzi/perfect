(ns perfect.poi
  (:require
   [clojure.java.io :as io]
   [com.rpl.specter :refer :all]
   [clojure.set :as st]
   [clojure.java.data :as jd]
   [perfect.utils :as utils :reload true]
   [camel-snake-kebab.core :as csk]
   [com.evocomputing.colors :as colors]
   [perfect.poi.full :as full :reload true]
   [perfect.poi.short :as shrt :reload true]
   [perfect.poi.fast :as fast :reload true])

  (:import
   (org.apache.poi.xssf.usermodel XSSFWorkbook
                                  XSSFCellStyle
                                  DefaultIndexedColorMap
                                  DefaultIndexedColorMap XSSFColor)
   (org.apache.poi.ss.usermodel Cell
                                CellType
                                Row
                                Sheet
                                Workbook
                                DateUtil
                                Row$MissingCellPolicy
                                CellStyle
                                HorizontalAlignment
                                VerticalAlignment
                                FillPatternType)
   (org.apache.poi.ss.util WorkbookUtil)
   (java.lang String Number Boolean)
   (clojure.lang Keyword)
   (java.util Date)))

(set! *warn-on-reflection* true)

(def ^:dynamic *row-missing-policy* Row$MissingCellPolicy/CREATE_NULL_AS_BLANK)

(def cache (atom {}))

(defn excel-file ^java.io.InputStream
  [path]
  (io/input-stream path))

(defprotocol XCell
  (^Cell create-cell [^Row row col] "Given a row create a cell at the given position")
  (set-cell [^Cell cell] [x ^Cell cell] [x ^Row row col] "Set the value of a cell"))

(extend-protocol XCell
  Row
  (create-cell [^Row row col] (.createCell row (int col)))

  Cell
  (set-cell [^Cell cell] (.setCellType cell CellType/BLANK))

  String
  (set-cell
    ([^String s ^Cell cell] (.setCellValue cell s))
    ([^String s ^Row row col] (set-cell s (create-cell row col))))

  Number
  (set-cell
    ([^Number n ^Cell cell] (.setCellValue cell (double n)))
    ([^Number n ^Row row col] (set-cell n (create-cell row col))))

  Boolean
  (set-cell
    ([^Boolean b ^Cell cell] (.setCellValue cell b))
    ([^Boolean b ^Row row col] (set-cell b (create-cell row col))))

  Keyword
  (set-cell
    ([^Keyword k ^Cell cell] (.setCellValue cell (name k)))
    ([^Keyword k ^Row row col] (set-cell k (create-cell row col))))

  Date
  (set-cell
    ([^Date d ^Cell cell] (.setCellValue cell d))
    ([^Date d ^Row row col] (set-cell d (create-cell row col))))

  nil
  (set-cell
    ([null ^Cell cell] (set-cell cell))
    ([null ^Row row col] (set-cell (create-cell row col)))))

(defprotocol XRow
  (^Row create-row [^Sheet sheet row-n] [^Sheet sheet row-n value col]
   "Create a row in the given sheet at the given position"))

(extend-protocol XRow
  Sheet
  (create-row
    ([^Sheet sheet row-n] (.createRow sheet (int row-n)))
    ([^Sheet sheet row-n value col]
     (let [row (create-row sheet row-n)
           _   (set-cell value row col)]
       row))))

(defprotocol XSheet
  (^Sheet create-sheet [^Workbook workbook] [^Workbook workbook ^String sheetname]
   "Create a new sheet in the given workbook"))

(extend-protocol XSheet
  Workbook
  (create-sheet
    ([^Workbook workbook] (.createSheet workbook))
    ([^Workbook workbook ^String sheetname]
     (.createSheet workbook (WorkbookUtil/createSafeSheetName sheetname)))))

(comment (defmethod jd/from-java XSSFCellStyle
           [^CellStyle cs]
           (let [data {:h-alignment (h-alignment cs)
                       :v-alignment (v-alignment cs)
                       :data-format (cell-format cs)
                       :wrapped?    (.getWrapText cs)
                       :background  (.getFillBackgroundColorColor cs)
                       :foreground  (get-color cs)
                       :pattern     (.getFillPattern cs)}]
             data)))

(defn enum->key
  [enum]
  (csk/->kebab-case-keyword (str enum)))

(defn h-alignment
  ([^CellStyle cs]
   (enum->key (.getAlignment cs)))
  ([^CellStyle cs style-keyword]
   (.setAlignment cs
                  (HorizontalAlignment/forInt
                   (style-keyword utils/horizontal-alignment)))))

(defn v-alignment
  [^CellStyle cs]
  (enum->key (.getVerticalAlignment cs)))

(defn cell-format
  [^CellStyle cs]
  (enum->key (.getDataFormatString cs)))

(defn get-color
  [^CellStyle cs]
  (let [rgb (.getRGB (.getFillForegroundColorColor cs))]
    (map #(+ % 0xff) rgb)))

(defn create-workbook
  ([]
   (jd/from-java (XSSFWorkbook.))))

(defn read-workbook
  ([path]
   (with-open [ef (excel-file path)]
     (full/from-java (XSSFWorkbook. ef))))
  ([path method]
   (with-open [ef (excel-file path)]
     (let [ms     {:full  full/from-java
                   :short full/from-java
                   :fast  full/from-java}
           method (method ms)]
       (method (XSSFWorkbook. ef))))))

(defn cell?
  [obj-m]
  (if (= :cell (:level obj-m))
    true
    false))

(defn blank?
  [obj]
  (let [t (:type obj)]
    (if t
      (= "BLANK" t)
      (empty? (:objs obj)))))
 
(defn clean-cell
  [wb-map]
  (setval [:objs ALL :objs ALL :objs ALL :objs] NONE wb-map))

; Styling stuff

(defn style
  [^Cell cell]
  (let [stl (.getCellStyle cell)]
    (jd/from-java stl)))

(defn map-set-cells
  [row values]
  (doall (map-indexed #(set-cell %2 row %1) values)))

(defn ordered-vals
  [header values]
  ((apply juxt header) values))

(comment

  (do
    (require '[criterium.core :refer [quick-bench
                                      bench
                                      with-progress-reporting]])
    (println)
    (println)
    (println "Full")
    (with-progress-reporting (quick-bench (read-workbook "resources/prova.xlsx")))
    (println)
    (println "Short")
    (with-progress-reporting
      (quick-bench
       (shrt/read-short "prova.xlsx"))))


  (let [wb    (create-workbook "prova.xlsx")
        sheet (first (sheets wb))
        cls   (cells sheet)]
    (with-open [o (clojure.java.io/output-stream "prova2.xlsx")]
      (.write wb o)))

  (let [wb    (create-workbook)
        sheet (create-sheet wb)
        row   (create-row sheet 0)
        cell  (create-cell row 0)
        stl   (.createCellStyle wb)
        cl    (.setFillPattern stl FillPatternType/SOLID_FOREGROUND)
        cl    (.setFillForegroundColor stl (short 15))
        _     (set-cell 1 cell)
        cell  (.setCellStyle cell stl)]
    (with-open [o (clojure.java.io/output-stream "colore.xlsx")] 
      (.write wb o))))

; TODO it makes sense to reason in a column format for config data and for writing/reading purposes, the only issue is that POI reasons in a row-based format

; TODO provare la libreria di joinr
