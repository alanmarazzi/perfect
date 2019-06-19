(ns perfect.reader
  (:require
   [perfect.reader.full  :as full :reload true]
   [perfect.reader.table :as tbl :reload true]
   [perfect.reader.fast  :as fast :reload true]))

(set! *warn-on-reflection* true)

(defn read-workbook
  "Read an Excel file to get a Clojure data
  representation. There are 3 reading options
  and a way to get a single sheet from a
  workbook.

  ```
  (read-workbook \"myexcelfile.xlsx\")
  (read-workbook \"myexcelfile.xlsx\" :method :fast)
  (read-workbook \"myexcelfile.xlsx\" :method :table :sheetid 0)
  (read-workbook \"myexcelfile.xlsx\" :method :full :sheetid \"Mysheet\")
  ```"
  [path & {:keys [method sheetid]
           :or   {method :full}}]
  (let [opts {:full  full/read-workbook
              :table tbl/read-workbook
              :fast  fast/read-workbook}
        f    (method opts)]
    (if sheetid
      (f path sheetid)
      (f path))))
