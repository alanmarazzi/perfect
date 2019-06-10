(ns perfect.poi.short
  (:import
   (org.dhatim.fastexcel.reader ReadableWorkbook
                                Sheet
                                Row
                                Cell
                                CellType)))

(defn cell-value
  ([^Cell cell] (cell-value cell (.getType cell)))
  ([^Cell cell cell-type]
   (condp = cell-type
     CellType/EMPTY nil
     CellType/STRING (.asString cell)
     CellType/NUMBER (.asNumber cell)
     CellType/BOOLEAN (.asBoolean cell)
     CellType/FORMULA {:formula (.getFormula cell)
                       :value   (.getValue cell)}
     CellType/ERROR {:error (.getValue cell)}
     :unsupported)))

(defn sheets
  [wb]
  (-> (.getSheets wb) .iterator iterator-seq))

(defn sheet
  [wb sheet]
  (cond
    (number? sheet) (try (.. wb (getSheet sheet) get)
                         (catch java.util.NoSuchElementException e
                           (throw (ex-info "" {:cause :this} e))))
    (string? sheet) (-> wb (.findSheet sheet))))

(defn rows
  [sheet]
  (-> (.openStream sheet) .iterator iterator-seq))

(defn cells
  [row]
  (-> (.iterator row) iterator-seq))

(defprotocol POIMapper
  (from-java [obj] [obj opt]))

(extend-protocol POIMapper
  nil
  (from-java [null]
    nil)

  Cell
  (from-java [^Cell cell]
    (cell-value cell))

  Sheet
  (from-java [^Sheet sheet]
    (into []
          (comp (map from-java)
                (filter #(not-every? nil? %)))
          (rows sheet)))

  Row
  (from-java [^Row row]
    (mapv from-java (cells row)))

  ReadableWorkbook
  (from-java
    ([^ReadableWorkbook wb]
     (map from-java (sheets wb)))
    ([^ReadableWorkbook wb sheetn]
     (from-java (sheet wb sheetn)))))

(defn read-workbook
  ([path]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-java (ReadableWorkbook. ef))))
  ([path opt]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-java (ReadableWorkbook. ef) opt))))
