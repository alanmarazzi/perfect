(defproject aida "0.1.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "EPL-2.0 OR GPL-2.0-or-later WITH Classpath-exception-2.0"
            :url  "https://www.eclipse.org/legal/epl-2.0/"}
  :dependencies [[org.clojure/clojure "1.10.0"]
                 [org.clojure/core.async "0.4.490"]
                 [com.climate/claypoole "1.1.4"]
                 [camel-snake-kebab "0.4.0"]
                 [com.evocomputing/colors "1.0.4"]
                 [simple-time "0.2.0"]
                 [org.apache.poi/poi "4.1.0"]
                 [org.apache.poi/poi-ooxml "4.1.0"]
                 [org.clojure/core.cache "0.7.2"]
                 [com.rpl/specter "1.1.2"]
                 [org.clojure/java.data "0.1.1"]
                 [org.dhatim/fastexcel-reader "0.10.2"]
                 [techascent/tech.datatype "4.0-alpha18"]

                 [spork "0.2.1.1-SNAPSHOT"]]
  :repl-options {:init-ns perfect.poi}
  ; :profiles {:dev {:dependencies [[org.openjfx/javafx-fxml     "11.0.1"]
  ;                                 [org.openjfx/javafx-controls "11.0.1"]
  ;                                 [org.openjfx/javafx-swing    "11.0.1"]
  ;                                 [org.openjfx/javafx-base     "11.0.1"]
  ;                                 [org.openjfx/javafx-web      "11.0.1"]]}}
  ; :resource-paths ["resources/REBL-0.9.157.jar"]
  )
