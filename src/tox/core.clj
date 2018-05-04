(ns tox.core
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:require [clojure.tools.cli :refer [parse-opts]]
            [clojure.java.jdbc :as j]
            [clojure.string :as str]
            [instaparse.core :as insta])
  (:gen-class))


;; (def db {:classname  "oracle.jdbc.OracleDriver"
;;          :subprotocol    "oracle:thin"
;;          :subname        "127.0.0.1:1521:orcl" 
;;          :user               "alan"
;;          :password       "alan"})
(def db (:db (read-string (slurp "db.conf"))))

;;(time (j/query db ["select sysdate from dual"]))

;; (def cells (j/query db
;;   ["select object_name oname,object_id,created,object_type from all_objects where rownum<10"]
;;   {:as-arrays? true :keywordize? false}))

;; (let [wb (create-workbook "objects"
;;                           cells
;;                           )]
;;    (save-workbook! "objects.xlsx" wb))

(def parser
  (insta/parser "S = <'('> members <')'>
                 <members> = member (<','> members)*
                 <member> = A | AB
                 A = #'[^\\(\\)\\,]+'
                 B = <'('> A <','> A <')'>
                 AB = A B"))

(def cli-options
  ;; An option with a required argument
  [
   ;; ["-p" "--port PORT" "Port number"
   ;;  :default 80
   ;;  :parse-fn #(Integer/parseInt %)
   ;;  :validate [#(< 0 % 0x10000) "Must be a number between 0 and 65536"]]
   ;; ;; A non-idempotent option
   ;; ["-v" nil "Verbosity level"
   ;;  :id :verbosity
   ;;  :default 0
   ;;  :assoc-fn (fn [m k _] (update-in m [k] inc))
   ;;  ]
   ;; tox parameters
   ["-i" "--infile filename" "q2x: template excel file.\nx2d: data excel&txt file."
    :validate [#(re-find #"\.(xls|csv|txt)" %) "input file must be .xls|.xlsx|.csv|.txt file."]]
   ["-o" "--outfile filename" "output excel&txt file."
    :validate [#(re-find #"\.(xls|csv|txt)" %) "output file must be .xls|.xlsx|.csv|.txt file."]]
   ["-d" "--delimiter delimiter" "field delimiter."
    :default ","]
   ["-q" "--qualifier qualifier" "field qualifier."
    :default "\""]
   ["-s" "--sql script" "sql script file"
    :validate [#(re-find #"\.sql" %) "Sql script file must be .sql file."]]
   ["-w" "--sheet workstheets" "worksheet names. exp:(sheet1,sheet2) or (sheet1(loc1,loc2),sheet1(loc3,loc4))"
    :parse-fn (fn [s] (->> (parser s)
                           (insta/transform {:S vector :A str :B hash-map :AB hash-map})))
    :validate [#(not (vector? (first %))) "sheet must be (s1,s2(a1,a2)..)" ]
   ]
   ;; A boolean option defaulting to nil
   ["-h" "--help"]])

(defn usage [options-summary]
  (->> ["数据库报表小工具tox。实现数据库与excel报表或txt文件之间数据互传。"
	""
	"Usage: program-name(tox) [options] action"
	""
	"Options:"
	options-summary
	""
	"Actions:"
	"  q2x    query to excel file."
	"  q2c    query to csv file(s)."
	"  x2d    excel file to db."
	""
	"Please refer to the manual page for more information."]
       (str/join \newline)))

(defn error-msg [errors]
  (str "The following errors occurred while parsing your command:\n\n"
       (str/join \newline errors)))

(defn validate-args
  "Validate command line arguments. Either return a map indicating the program
  should exit (with a error message, and optional ok status), or a map
  indicating the action the program should take and the options provided."
  [args]
  (let [{:keys [options arguments errors summary]} (parse-opts args cli-options)]
    (cond
      (:help options) ; help => exit OK with usage summary
      {:exit-message (usage summary) :ok? true}
      errors ; errors => exit with description of errors
      {:exit-message (error-msg errors)}
      ;; custom validation on arguments
      (and (= 1 (count arguments))
           (#{"q2x" "q2c" "x2d"} (first arguments)))
      {:action (first arguments) :options options}
      :else ; failed custom validation => exit with usage summary
      {:exit-message (usage summary)})))

(defn exit [status msg]
  (println msg)
  (System/exit status)
  )

(defn get-sqls [opts]
  (spit))
(defn q2x! [opts]
	opts)

(defn q2c! [opts]
	opts)

(defn x2d! [opts]
	opts)

(defn -main [& args]
  (let [{:keys [action options exit-message ok?]} (validate-args args)]
    (if exit-message
      (exit (if ok? 0 1) exit-message)
      (case action
        "q2x" (q2x! options)
        "q2c" (q2c! options)
        "x2d" (x2d! options)))))
