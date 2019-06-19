(ns perfect.reader.generics
  (:require
   [clojure.spec.alpha :as s]
   [expound.alpha :as expound]
   [expound.specs :as exspec])
  (:import
   (java.lang IllegalArgumentException)
   (org.apache.poi.ss.util WorkbookUtil)
   (org.apache.poi.ss.usermodel Workbook
                                Sheet
                                Row
                                Cell
                                CellType
                                DateUtil)))

(set! s/*explain-out* expound/printer)
(s/check-asserts true)

(defn valid-name?
  [nm]
  (try
    (do
      (WorkbookUtil/validateSheetName nm)
      true)
    (catch IllegalArgumentException e
      false)))

(expound/def ::valid-name?
  (s/and string? valid-name?)
  "should be an XLSX valid name: https://poi.apache.org/apidocs/4.1/org/apache/poi/ss/util/WorkbookUtil.html#createSafeSheetName-java.lang.String-")

(s/def ::sheet-identity (s/or :idx  ::exspec/nat-int
                              :name ::valid-name?))

(defn sheets
  [^Workbook wb]
  (map #(.getSheetAt wb %) (range (.getNumberOfSheets wb))))

(defn rows
  [^Sheet sheet]
  (seq sheet))

(defn cells
  [row]
  (seq row))

(defn columnar
  [d header?]
  (if header?
    (zipmap (first d) (apply mapv vector (rest d)))
    (apply mapv vector (rest d))))

(defn blank?
  [cell]
  (= :blank (:type cell)))
