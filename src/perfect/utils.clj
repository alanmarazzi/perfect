(ns perfect.utils)

(defn slurp-bytes
  "Slurp the bytes from a slurpable thing"
  [x]
  (with-open [out (java.io.ByteArrayOutputStream.)]
    (clojure.java.io/copy (clojure.java.io/input-stream x) out)
    (.toByteArray out)))

(comment
  (defn write-dataset-edn! [out-file raw-dataset-map]
    (with-open [w (clojure.java.io/writer out-file)]
      (binding [*out* w]
        (clojure.pprint/write raw-dataset-map)))))

(def data-formats {:general             0
                   :number              1
                   :decimal             2
                   :comma               3
                   :accounting          4
                   :dollars             5
                   :red-neg             6
                   :cents               7
                   :dollars-red-neg     8
                   :percentage          9
                   :decimal-percentage  10
                   :scientific-notation 11
                   :short-ratio         12
                   :ratio               13
                   :date                14
                   :day-month-year      15
                   :day-month-name      16
                   :month-name-year     17
                   :hour-am-pm          18
                   :time-am-pm          1
                   :hour                20
                   :time                21
                   :datetime            22})

(def color-indices {:black                 8
                    :violet                20
                    :blue-grey             54
                    :dark-yellow           19
                    :automatic             64
                    :grey-50-percent       23
                    :light-cornflower-blue 31
                    :light-turquoise       41
                    :white                 9
                    :plum                  61
                    :orange                53
                    :teal                  21
                    :olive-green           59
                    :red                   10
                    :lavender              46
                    :light-orange          52
                    :brown                 60
                    :light-green           42
                    :light-yellow          43
                    :royal-blue            30
                    :gold                  51
                    :aqua                  49
                    :coral                 29
                    :light-blue            48
                    :blue                  12
                    :cornflower-blue       24
                    :dark-green            58
                    :grey-25-percent       22
                    :green                 17
                    :lemon-chiffon         26
                    :lime                  50
                    :dark-blue             18
                    :sky-blue              40
                    :pale-blue             44
                    :dark-red              16
                    :dark-teal             56
                    :grey-80-percent       63
                    :indigo                62
                    :sea-green             57
                    :pink                  14
                    :turquoise             15
                    :tan                   47
                    :yellow                13
                    :orchid                28
                    :grey-40-percent       55
                    :rose                  45
                    :maroon                25
                    :bright-green          11})

(def underline-indices
  {:none 0 :single 1 :double 2 :single-accounting 33 :double-accounting 34})

(def horizontal-alignment
  {:general          0
   :left             1
   :center           2
   :right            3
   :fill             4
   :justify          5
   :center-selection 6
   :distributed      7})

(def vertical-alignment
  {:top         0
   :center      1
   :bottom      2
   :justify     3
   :distributed 4})