(ns perfect.poi.fast
  (:require
    [perfect.poi.generics :as generics]
    [net.danielcompton.defn-spec-alpha :as ds])
  (:import
    (org.dhatim.fastexcel.reader ReadableWorkbook
                                 Sheet
                                 Row
                                 Cell
                                 CellType)))

(defn cell-value
  "Return proper getter based on cell-value"
  ([^Cell cell] (cell-value cell (.getType cell)))
  ([^Cell cell cell-type]
   (condp = cell-type
     CellType/EMPTY nil
     CellType/STRING (.asString cell)
     CellType/NUMBER (float (.asNumber cell))
     CellType/BOOLEAN (.asBoolean cell)
     CellType/FORMULA {:formula (.getFormula cell)}
     CellType/ERROR {:error (.getValue cell)}
     :unsupported)))

(defn cell-type
  [cell]
  (condp = cell
    CellType/EMPTY :blank
    CellType/STRING :str
    CellType/NUMBER :numeric
    CellType/BOOLEAN :bool
    CellType/FORMULA :formula
    CellType/ERROR :error
    :unsupported))

(defn sheets
  [wb]
  (iterator-seq (.. wb getSheets iterator)))

(ds/defn sheet
  [wb
   sheet :- ::generics/sheet-identity]
  (s/assert ::generics/sheet-identity sheet)
  (cond
    (pos-int? sheet) (.. wb (getSheet sheet) get)
    (string? sheet)  (.. wb (findSheet sheet) get)))

(defn rows
  [sheet]
  (iterator-seq (.. sheet openStream iterator)))

(defn cells
  [row]
  (iterator-seq (.iterator row)))

(defprotocol POIMapper
  (from-java [obj] [obj opt]))

(extend-protocol POIMapper
  Cell
  (from-java [^Cell cell]
    {:type   (cell-type (.getType cell))
     :row-id (inc (.getRow (.getAddress cell)))
     :col-id (inc (.getColumn (.getAddress cell)))
     :level  :cell
     :value  (cell-value cell)})

  Sheet
  (from-java [^Sheet sheet]
    {:sheet-name (.getName sheet)
     :sheet-id   (inc (.getIndex sheet))
     :rows       (into []
                       (comp
                         (map from-java)
                         (filter #(not-empty (:cells %))))
                       (rows sheet))
     :level      :sheet})

  Row
  (from-java [^Row row]
    {:n-cells    (.getCellCount row)
     :row-number (.getRowNum row)
     :cells      (into []
                       (comp
                         (map from-java)
                         (filter (complement generics/blank?))
                         (keep identity))
                       (cells row))})

  ReadableWorkbook
  (from-java
    ([^ReadableWorkbook wb]
     (let [s (into [] (map from-java) (sheets wb))
           c (count s)]
       {:sheet-count c
        :sheets      s
        :level       :workbook}))
    ([^ReadableWorkbook wb sheetn]
     (let [s (into [] (map from-java) (sheets wb))
           c (count s)]
       {:sheet-count c
        :sheets      (from-java (sheet wb sheetn))
        :level       :workbook})))

  nil
  (from-java
    [obj]
    nil))

(defn read-workbook
  ([path]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-java (ReadableWorkbook. ef))))
  ([path opt]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-java (ReadableWorkbook. ef) opt))))
