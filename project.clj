(defproject perfect "0.0.1-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "EPL-2.0 OR GPL-2.0-or-later WITH Classpath-exception-2.0"
            :url  "https://www.eclipse.org/legal/epl-2.0/"}
  :dependencies [[org.clojure/clojure "1.10.1"]
                 [org.apache.poi/poi "4.1.0"]
                 [org.apache.poi/poi-ooxml "4.1.0"]
                 [org.dhatim/fastexcel-reader "0.10.2"]
                 [org.clojure/spec.alpha "0.2.176"]
                 [net.danielcompton/defn-spec-alpha "0.1.0"]
                 [expound "0.7.2"]]
  :profiles {:dev {:dependencies [[org.clojure/test.check "0.10.0-alpha4"]]}})