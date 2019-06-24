(ns perfect.benchmark
  "Temporary ns for benchmarking purposes"
  (:require
   [perfect.reader.fast :as fast :reload true]
   [perfect.reader.generics :as generics]
   [criterium.core :refer [bench]]
   [clojure.core.async :as a])
  (:import
   (org.dhatim.fastexcel.reader ReadableWorkbook
                                Sheet
                                Row
                                Cell
                                CellType)))

(def benchfile "resources/test.xlsx")

(defn bench-fast
  []
  (println "\nfast result")
  (bench (fast/read-workbook benchfile)))

(defprotocol RecordMapper
  (from-excel [obj]))

(defrecord xlcell  [type row-id col-id value])
(defrecord xlrow   [n-cells row-id cells])
(defrecord xlsheet [sheet-name sheet-id rows])
(defrecord xlworkbook [sheet-count sheets])

(extend-protocol RecordMapper
  Cell
  (from-excel [^Cell cell]
    (->xlcell (fast/cell-type (.getType cell))
              (.getRow (.getAddress cell))
              (.getColumn (.getAddress cell))
              (fast/cell-value cell)))
  Row
  (from-excel [^Row row]
    (->xlrow (.getCellCount row)
             (dec (.getRowNum row))
             (into []
                   (comp
                    (map from-excel)
                    (filter (complement generics/blank?))
                    (keep identity))
                   (fast/cells row))))

  Sheet
  (from-excel [^Sheet sheet]
    (->xlsheet (.getName sheet)
               (.getIndex sheet)
               (into []
                     (comp
                      (map from-excel)
                      (filter #(not-empty (:cells %))))
                     (fast/rows sheet))))

  ReadableWorkbook
  (from-excel [^ReadableWorkbook wb]
    (let [s (into [] (map from-excel) (fast/sheets wb))
          c (count s)]
      (->xlworkbook c s)))

  nil
  (from-excel
    [obj]
    nil))

(defn bench-record
  []
  (println "\n records result")
  (bench (with-open [ef (clojure.java.io/input-stream benchfile)]
           (from-excel (ReadableWorkbook. ef)))))

(defn ireducer [^java.util.Iterator iter]
  (reify
    java.lang.Iterable
    (iterator [this] iter)
    clojure.core.protocols/CollReduce
    (coll-reduce [this f]
      (when (.hasNext iter)
        (clojure.core.protocols/coll-reduce this f (.next iter))))
    (coll-reduce [this f init]
      (loop [acc init]
        (cond (reduced? acc)
              @acc
              (.hasNext iter)
              (recur (f acc (.next iter)))
              :else acc)))
    clojure.lang.Seqable
    (seq [this]
      (iterator-seq iter))))

(defn sheets
  [^ReadableWorkbook wb]
  (iterator-seq (.. wb getSheets iterator)))

(defn rows
  [^Sheet sheet]
  (ireducer (.. sheet openStream iterator)))

(defn cells
  [^Row row]
  (ireducer  (.iterator row)))

(defn ->cell [^Cell cell]
  (when cell
    {:type   (fast/cell-type (.getType cell))
     :row-id (.getRow (.getAddress cell))
     :col-id (.getColumn (.getAddress cell))
     :value  (fast/cell-value cell)}))

(defn ->cell-record [^Cell cell]
  (when cell
    (xlcell.  (fast/cell-type  (.getType cell)) 
              (.getRow    (.getAddress cell))
              (.getColumn (.getAddress cell))
              (fast/cell-value cell))))

(defn ->row [^Row row]
  (xlrow.  (.getCellCount row)
           (unchecked-dec (.getRowNum row))
           (->> (cells row)
                (reduce  (fn [^clojure.lang.PersistentVector$TransientVector acc v]
                           (let [v (->cell v)]
                             (if (and v (not (identical? :blank v)))
                               (.conj acc v)
                               acc)))
                         (transient []))
                persistent!)))

(defn ->row-record [^Row row]
  (xlrow.  (.getCellCount row)
           (unchecked-dec (.getRowNum row))
           (->> (cells row)
                (reduce  (fn [^clojure.lang.PersistentVector$TransientVector acc v]
                           (let [v (->cell-record v)]
                             (if (and v (not (identical? :blank v)))
                               (.conj acc v)
                               acc)))
                         (transient []))
                persistent!)))

(defn ->sheet [^Sheet sheet]
  {:sheet-name (.getName sheet)
   :sheet-id   (.getIndex sheet)
   :rows       (into []
                     (comp
                      (map ->row)
                      (filter #(not-empty (:cells %))))
                     (rows sheet))})

(defn ->sheet-record [^Sheet sheet]
  (xlsheet. (.getName sheet)
            (.getIndex sheet)
            (into []
                  (comp
                   (map ->row-record)
                   (filter #(not-empty (:cells %))))
                  (rows sheet))))

(defn ->book [^ReadableWorkbook wb]
  (let [s (into [] (map ->sheet) (sheets wb))
        c (count s)]
    {:sheet-count c
     :sheets      s}))

(defn ->book-record [^ReadableWorkbook wb]
  (let [s (into [] (map ->sheet-record) (sheets wb))
        c (count s)]
    (xlworkbook. c s)))

(defn bench-reducer
  []
  (println "\n reducer result")
  (bench (with-open [ef (clojure.java.io/input-stream benchfile)]
           (->book (ReadableWorkbook. ef)))))

(defn bench-reducer-record
  []
  (println "\n reducer-record result")
  (bench (with-open [ef (clojure.java.io/input-stream benchfile)]
           (->book-record (ReadableWorkbook. ef)))))

(defn -main
  []
  (bench-fast)
  (bench-record)
  (bench-reducer)
  (bench-reducer-record))
