(ns perfect.reader.full
  (:require
   [perfect.poi.generics :as generics :reload true]
   [net.danielcompton.defn-spec-alpha :as ds]
   [clojure.spec.alpha :as s])
  (:import
   (org.apache.poi.xssf.usermodel XSSFWorkbook)
   (org.apache.poi.ss.usermodel Cell
                                CellType
                                DateUtil
                                Row
                                Sheet
                                Workbook)))

(defn cell-value
  "Return the value of a cell using the proper method"
  ([^Cell cell] (cell-value cell (.getCellType cell)))
  ([^Cell cell cell-type]
   (condp = cell-type
     CellType/BLANK   nil
     CellType/STRING  (.getStringCellValue cell)
     CellType/NUMERIC (if (DateUtil/isCellDateFormatted cell)
                        (.getDateCellValue cell)
                        (.getNumericCellValue cell))
     CellType/BOOLEAN (.getBooleanCellValue cell)
     CellType/FORMULA {:formula (.getCellFormula cell)
                       :value   (cell-value cell (.getCachedFormulaResultType cell))}
     CellType/ERROR   {:error (.getErrorCellValue cell)}
     :unsupported)))

(ds/defn sheet
  "Get a sheet from a Workbook requiring it by id
  or by name"
  [wb
   sheetn :- ::generics/sheet-identity]
  (s/assert ::generics/sheet-identity sheetn)
  (cond
    (pos-int? sheetn) (.getSheetAt wb (dec sheetn))
    (string? sheetn)  (.getSheet wb sheetn)))

(defprotocol FullMapper
  "An Excel workbook is represented as a tree:

  Workbook
  └──Sheet
     └──Row
        └──Cell
           └──5

  The one above is a spreadsheet with one sheet and only one
  populated cell with the value 5. *Workbooks*, *Sheets* and
  *Rows* are pure abstractions: if there isn't any *Cell*
  with some value (even blank is ok) they don't exist.
  
  `FullMapper` dispatches the right method calls automatically 
  and doesn't require the user to make strictly ordered calls 
  to different functions.

  `from-excel` can be used directly by the user, but most of
  the time that's not what you want and you can just consider
  `FullMapper` as a *backend*, this makes possible to improve,
  change and add functionality without compromising compatibility."
  (from-excel [obj] [obj opt]))

(extend-protocol FullMapper
  Workbook
  (from-excel
    ([^Workbook wb]
     {:sheet-number (.getNumberOfSheets wb)
      :sheets       (into [] (map from-excel) (generics/sheets wb))})
    ([^Workbook wb sheetn]
     {:sheet-number (.getNumberOfSheets wb)
      :sheets       (from-excel (sheet wb sheetn))}))

  Sheet
  (from-excel [^Sheet sheet]
    {:first-row  (.getFirstRowNum sheet)
     :last-row   (.getLastRowNum sheet)
     :sheet-name (.getSheetName sheet)
     :first-col  (.getLeftCol sheet)
     :rows       (into []
                       (comp
                         (map from-excel)
                         (filter #(not-empty (:cells %))))
                       (generics/rows sheet))})

  Row
  (from-excel [^Row row]
    {:n-cells    (.getPhysicalNumberOfCells row)
     :row-number (.getRowNum row)
     :first-col  (.getFirstCellNum row)
     :last-col   (.getLastCellNum row)
     :cells      (into []
                       (comp
                         (map from-excel)
                         (keep identity))
                       (generics/cells row))})

  Cell
  (from-excel [^Cell cell]
    (let [cv (cell-value cell)]
      (when cv
        {:type    (str (.getCellType cell))
         :col-idx (.getColumnIndex cell)
         :value   cv})))

  nil
  (from-excel [obj]
    nil))

(defn read-workbook
  ([path]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-excel (XSSFWorkbook. ef))))
  ([path opt]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-excel (XSSFWorkbook. ef) opt))))