(ns perfect.reader.table
  (:import
   (org.dhatim.fastexcel.reader ReadableWorkbook
                                Sheet
                                Row
                                Cell
                                CellType)))

(defn cell-value
  "Get the value of a *Cell* calling the proper getter
  based on its type"
  ([^Cell cell] (cell-value cell (.getType cell)))
  ([^Cell cell cell-type]
   (condp = cell-type
     CellType/EMPTY   nil
     CellType/STRING  (.asString cell)
     CellType/NUMBER  (.asNumber cell)
     CellType/BOOLEAN (.asBoolean cell)
     CellType/FORMULA {:formula (.getFormula cell)
                       :value   (.getValue cell)}
     CellType/ERROR   {:error (.getValue cell)}
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

(defprotocol ShortMapper
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
  
  `ShortMapper` dispatches the right method calls automatically 
  and doesn't require the user to make strictly ordered calls 
  to different functions.

  `from-excel` can be used directly by the user, but most of
  the time that's not what you want and you can just consider
  `ShortMapper` as a *backend*, this makes possible to improve,
  change and add functionality without compromising compatibility."
  (from-excel [obj] [obj opt]))

(extend-protocol ShortMapper
  nil
  (from-excel [null]
    nil)

  Cell
  (from-excel [^Cell cell]
    (cell-value cell))

  Sheet
  (from-excel [^Sheet sheet]
    (into []
          (comp (map from-excel)
                (filter #(not-every? nil? %)))
          (rows sheet)))

  Row
  (from-excel [^Row row]
    (into [] (map from-excel) (cells row)))

  ReadableWorkbook
  (from-excel
    ([^ReadableWorkbook wb]
     (into [] (map from-excel) (sheets wb)))
    ([^ReadableWorkbook wb sheetn]
     (from-excel (sheet wb sheetn)))))

(defn read-workbook
  ([path]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-excel (ReadableWorkbook. ef))))
  ([path opt]
   (with-open [ef (clojure.java.io/input-stream path)]
     (from-excel (ReadableWorkbook. ef) opt))))
