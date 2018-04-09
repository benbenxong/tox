(ns tox.core
  (:require [clojure.tools.cli :refer [parse-opts]]
            [clojure.java.jdbc :as j]
            [clojure.string :as str]
            [instaparse.core :as insta])
  (:gen-class))

(def db {:classname  "oracle.jdbc.OracleDriver"
         :subprotocol    "oracle:thin"
         :subname        "127.0.0.1:1521:orcl" 
         :user               "alan"
         :password       "alan"})

;;(time (j/query db ["select sysdate from dual"]))

(def cli-options
  ;; An option with a required argument
  [["-p" "--port PORT" "Port number"
    :default 80
    :parse-fn #(Integer/parseInt %)
    :validate [#(< 0 % 0x10000) "Must be a number between 0 and 65536"]]
   ;; A non-idempotent option
   ["-v" nil "Verbosity level"
    :id :verbosity
    :default 0
    :assoc-fn (fn [m k _] (update-in m [k] inc))]
   ;; tox parameters
   ["-i" "--ifile filename" "input excel file"
   :validate [#(re-find #"\.xls" %) "Input excel file must be .xls|.xlsx file.] "]]
   ["-o" "--ofile filename" "output excel file"]
   ["-s" "--sql script" "sql script file"]
   ["-w" "--sheet workstheets" "worksheet names. exp:(sheet1,sheet2) or (sheet1(loc1,loc2),sheet1(loc3,loc4))"]
   ;; A boolean option defaulting to nil
   ["-h" "--help"]])
(def parser
            (insta/parser "S = <'('> members <')'>
                           <members> = member (<','> members)*
                           <member> = A | AB
                           A = #'[^\\(\\)\\,]+'
                           B = <'('> A <','> A <')'>
                           AB = A B"))

(->> (parser "(s1,s2,s(d1,顶戴枯))")
     (insta/transform {:S vector :A str :B hash-map :AB hash-map})) 


(defn -main
  [& args]
  (parse-opts args cli-options))
