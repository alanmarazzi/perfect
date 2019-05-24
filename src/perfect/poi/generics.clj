(ns perfect.poi.generics
  (:import
   (org.apache.poi.ss.usermodel Workbook
                                Sheet
                                Row
                                Cell
                                CellType
                                DateUtil)))

(defn cell-value
  "Return proper getter based on cell-value"
  ([^Cell cell] (cell-value cell (.getCellType cell)))
  ([^Cell cell cell-type]
   (condp = cell-type
     CellType/BLANK   nil
     CellType/STRING  (.getStringCellValue cell)
     CellType/NUMERIC (if (DateUtil/isCellDateFormatted cell)
                        (.getDateCellValue cell)
                        (.getNumericCellValue cell))
     CellType/BOOLEAN (.getBooleanCellValue cell)
     CellType/FORMULA {:formula (.getCellFormula cell)}
     CellType/ERROR   {:error (.getErrorCellValue cell)}
     :unsupported)))

(defn sheets
  [^Workbook wb] 
  (map #(.getSheetAt wb %) (range (.getNumberOfSheets wb))))

(defn rows
  [^Sheet sheet] 
  (seq sheet))

(defn cells
  [row]
  (seq row))

(defn values
  [row]
  (mapv cell-value (cells row)))
