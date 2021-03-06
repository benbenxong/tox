(defproject tox "0.1.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.8.0"]
                 [org.clojure/tools.cli "0.3.5"]
                 [clj-time "0.14.2"]
                 [clj-det-enc "1.0.0"]
                 [dk.ative/docjure "1.11.0"]
                 ;;[oracle.jdbc/oracledriver "11.2.0.1"]
                 [org.clojars.zentrope/ojdbc "11.2.0.3.0"]
                 [org.clojure/java.jdbc "0.7.5"]
                 [clojure.jdbc/clojure.jdbc-c3p0 "0.3.3"]
                 [instaparse "1.4.8"]
                 [mysql/mysql-connector-java "5.1.6"] 
                 [failjure "1.3.0"]]
  :main ^:skip-aot tox.core
  :target-path "target/%s"
  :resource-paths ["shared" "resources"]
  :profiles {:uberjar {:aot :all}})
