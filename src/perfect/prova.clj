(ns perfect.prova
  (:require
    [perfect.poi.full :as full :reload true]
    [perfect.poi.fast :as fast :reload true]
    [perfect.poi.generics :as gen :reload true]

    [clojure.spec.alpha :as s]
    [clojure.spec.gen.alpha :as sgen]
    [clojure.spec.test.alpha :as stest]
    [net.danielcompton.defn-spec-alpha :as ds]))

(def p (gen/columnar (fast/read-workbook "resources/pipbase.xlsx" 1) true))

(def t (gen/columnar (fast/read-workbook "resources/pipnew.xlsx" 1) true))

(defn n-ceduti
  [cols]
  (reduce + (keep identity (cols "nr ceduti"))))

(defn n-cedenti
  [cols]
  (count (keep identity (cols "CEDENTI"))))


(s/def ::big-even (s/and int? even? #(> % 1000)))
(s/valid? ::big-even :foo)
(s/valid? ::big-even 10)
(s/valid? ::big-even 100000000)

(s/def ::name-or-id (s/or :name string?
                          :id int?))
(s/valid? ::name-or-id "abc")
(s/valid? ::name-or-id 100)
(s/valid? ::name-or-id :f)
(s/explain ::name-or-id :f)

(def email-regex #"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,63}$")
(s/def ::email-type (s/and string? #(re-matches email-regex %)))

(s/def ::acctid int?)
(s/def ::first-name string?)
(s/def ::last-name string?)
(s/def ::email ::email-type)
(s/def ::phone string?)

(s/def ::person (s/keys :req [::first-name ::last-name ::email]
                        :opt [::phone]))
(s/valid? ::person
          {::first-name "Bugs"
           ::last-name  "Bunny"
           ::email      "bugs@example.com"})

(s/def ::ingredient (s/cat :quantity number? :unit keyword?))
(s/conform ::ingredient [2 :teaspoon])
(s/explain ::ingredient [11 "peaches"])

(defn person-name
  [person]
  {:pre  [(s/valid? ::person person)]
   :post [(s/valid? string? %)]}
  (str (::first-name person) " " (::last-name person)))

(defn person-name
  [person]
  (let [p (s/assert ::person person)]
    (str (::first-name p) " " (::last-name p))))

(defn ranged-rand
  "Returns random int in range start <= rand < end"
  [start end]
  (+ start (long (rand (- end start)))))

(s/fdef ranged-rand
        :args (s/and (s/cat :start int? :end int?)
                     #(< (:start %) (:end %)))
        :ret int?
        :fn (s/and #(>= (:ret %) (-> % :args :start))
                   #(< (:ret %) (-> % :args :end))))


(sgen/sample (s/gen (s/cat :k keyword? :ns (s/+ number?))))


(def suit? #{:club :diamond :heart :spade})
(def rank? (into #{:jack :queen :king :ace} (range 2 11)))
(def deck (for [suit suit? rank rank?] [rank suit]))

(s/def ::card (s/tuple rank? suit?))
(s/def ::hand (s/* ::card))

(s/def ::name string?)
(s/def ::score int?)
(s/def ::player (s/keys :req [::name ::score ::hand]))

(s/def ::players (s/* ::player))
(s/def ::deck (s/* ::card))
(s/def ::game (s/keys :req [::players ::deck]))

(sgen/generate (s/gen ::player))
(sgen/generate (s/gen ::game))

(s/exercise (s/cat :k keyword? :ns (s/+ number?)) 5)
(s/exercise-fn `ranged-rand)

(defn divisible-by [n] #(zero? (mod % n)))
(sgen/sample (s/gen (s/and int?
                           #(> % 0)
                           (divisible-by 3))))

(stest/instrument `ranged-rand)
(stest/check `ranged-rand)

(defn ranged-rand ;; BROKEN!
  "Returns random int in range start <= rand < end"
  [start end]
  (+ start (long (rand (- start end)))))
(stest/abbrev-result (first (stest/check `ranged-rand)))

(ds/defn adder :- int?
  [x :- int?]
  (inc x))
