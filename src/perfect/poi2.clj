(ns perfect.poi2
  (:require
   [clojure.java.io :as io]
   [com.rpl.specter :refer :all]
   [clojure.set :as st]
   [clojure.java.data :as jd]
   [aida.utils :as utils :reload true]
   [camel-snake-kebab.core :as csk]
   [com.evocomputing.colors :as colors]
   [clojure.zip :as zip])
  (:import 
   (org.dhatim.fastexcel.reader ReadableWorkbook
                                Sheet Row Cell
                                CellType)))

(defn excel-file ^java.io.InputStream
  [path]
  (io/input-stream path))

(defmethod jd/from-java ^ReadableWorkbook
  [^ReadableWorkbook wb]
  {})

(defn workbook
  [path]
  (ReadableWorkbook. (excel-file path)))

(defn sheet
  [path]
  (let [w (workbook (excel-file path))]
    (.getSheets w)))

(defn row
  [path])

(defn value
  [^Cell cell]
  (when (not (nil? cell))
    (.getValue cell)))



(def r (second (.read (sheet "prova.xlsx"))))

(first (map value (seq r)))