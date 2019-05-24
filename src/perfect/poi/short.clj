(ns perfect.poi.short
  (:import
   
   (org.dhatim.fastexcel.reader ReadableWorkbook
                                Sheet
                                Row
                                Cell
                                CellType)))

(comment (org.apache.poi.ss.usermodel Cell
                                      Row
                                      Sheet
                                      Workbook))

(defn cell-value
  "Return proper getter based on cell-value"
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

(defn sheets
  [wb]
  (-> (.getSheets wb) .iterator iterator-seq))

(defn rows
  [sheet]
  (-> (.openStream sheet) .iterator iterator-seq))

(defn cells
  [row]
  (-> (.iterator row) iterator-seq))

(defprotocol POIMapper
  (from-java [obj]))

(extend-protocol POIMapper
  Cell
  (from-java [^Cell cell]
    {;:type    (str (.getCellType cell))
     ;:row-idx (.getRowIndex cell)
                                        ;:col-idx (.getColumnIndex cell)
     :row-idx (.getRow (.getAddress cell))
     :col-idx (.getColumn (.getAddress cell))
     :level   :cell
     :value   (cell-value cell)})

  Sheet
  (from-java [^Sheet sheet]
    {;:first-row  (.getFirstRowNum sheet)
     ;:last-row   (.getLastRowNum sheet)
     ;:sheet-name (.getSheetName sheet)
                                        ;:first-col  (.getLeftCol sheet)
     :sheet-idx (.getIndex sheet)
     :cells       (map from-java (rows sheet)) 
     :level      :sheet})

  Row
  (from-java [^Row row]
   ;(into [] (map from-java) (seq row))
    (map from-java (cells row))
    )

  ReadableWorkbook
  (from-java [^ReadableWorkbook wb]
    (let [s (map from-java (sheets wb))
          c (count s)]
      {:sheet-number c ;(.getNumberOfSheets wb)
       :sheets        s                  ;(into [] (map from-java) (sheets wb))
       :level        :workbook})))

(defn read-short
  [path]
  (with-open [ef (clojure.java.io/input-stream path)]
    (from-java (ReadableWorkbook. ef))))
