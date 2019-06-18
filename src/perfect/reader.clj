(ns perfect.reader
  (:require
   [perfect.reader.full :as full :reload true]
   [perfect.reader.short :as shrt :reload true]
   [perfect.reader.fast :as fast :reload true]))

(set! *warn-on-reflection* true)

(defn read-workbook
  [path & {:keys [method sheetn]
           :or   {method :full}}]
  (let [opts {:full  full/read-workbook
              :short shrt/read-workbook
              :fast  fast/read-workbook}
        f    (method opts)]
    (if sheetn
      (f path sheetn)
      (f path))))