(ns perfect.reader.fast
  (:require
   [perfect.reader.generics :as generics]
   [clojure.spec.alpha :as s]
   [net.danielcompton.defn-spec-alpha :as ds])
  (:import
    (org.dhatim.fastexcel.reader ReadableWorkbook
                                 Sheet
                                 Row
                                 Cell
                                 CellType)))

(defn cell-value
  "Return the value of a cell using the proper method"
  ([^Cell cell] (cell-value cell (.getType cell)))
  ([^Cell cell cell-type]
   (condp = cell-type
     CellType/EMPTY   nil
     CellType/STRING  (.asString cell)
     CellType/NUMBER  (float (.asNumber cell))
     CellType/BOOLEAN (.asBoolean cell)
     CellType/FORMULA {:formula (.getFormula cell)}
     CellType/ERROR   {:error (.getValue cell)}
     :unsupported)))

(defn cell-type
  "From Excel types to keywords"
  [cell]
  (condp = cell
    CellType/EMPTY   :blank
    CellType/STRING  :str
    CellType/NUMBER  :numeric
    CellType/BOOLEAN :bool
    CellType/FORMULA :formula
    CellType/ERROR   :error
    :unsupported))

(defn sheets
  [wb]
  (iterator-seq (.. wb getSheets iterator)))

(ds/defn sheet
  [wb
   sheetid :- ::generics/sheet-identity]
  (s/assert ::generics/sheet-identity sheetid)
  (cond
    (number? sheet) (.. wb (getSheet sheetid) get)
    (string? sheet) (.. wb (findSheet sheetid) get)))

(defn rows
  [sheet]
  (iterator-seq (.. sheet openStream iterator)))

(defn cells
  [row]
  (iterator-seq (.iterator row)))

(defprotocol FastMapper
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

  `FastMapper` dispatches the right method calls automatically
  and doesn't require the user to make strictly ordered calls
  to different functions.

  `from-excel` can be used directly by the user, but most of
  the time that's not what you want and you can just consider
  `FastMapper` as a *backend*, this makes possible to improve,
  change and add functionality without compromising compatibility."
  (from-excel [obj] [obj opt]))

(extend-protocol FastMapper
  Cell
  (from-excel [^Cell cell]
    {:type   (cell-type (.getType cell))
     :row-id (.getRow (.getAddress cell))
     :col-id (.getColumn (.getAddress cell))
     :value  (cell-value cell)})

  Sheet
  (from-excel [^Sheet sheet]
    {:sheet-name (.getName sheet)
     :sheet-id   (.getIndex sheet)
     :rows       (into []
                       (comp
                        (map from-excel)
                        (filter #(not-empty (:cells %))))
                       (rows sheet))})

  Row
  (from-excel [^Row row]
    {:n-cells (.getCellCount row)
     :row-id  (dec (.getRowNum row))
     :cells   (into []
                    (comp
                     (map from-excel)
                     (filter (complement generics/blank?))
                     (keep identity))
                    (cells row))})

  ReadableWorkbook
  (from-excel
    ([^ReadableWorkbook wb]
     (let [s (into [] (map from-excel) (sheets wb))
           c (count s)]
       {:sheet-count c
        :sheets      s}))
    ([^ReadableWorkbook wb sheetn]
     (let [s (into [] (map from-excel) (sheets wb))
           c (count s)]
       {:sheet-count c
        :sheets      (from-excel (sheet wb sheetn))})))

  nil
  (from-excel
    [obj]
    nil))

(defn read-workbook
  ""
  ([path]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-excel (ReadableWorkbook. ef))))
  ([path opt]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-excel (ReadableWorkbook. ef) opt))))
