(ns tox.core
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:require [clojure.tools.cli :refer [parse-opts]]
            [clojure.java.jdbc :as j]
            [clojure.string :as str]
            [instaparse.core :as insta]
            [det-enc.core :as det]
            [clojure.java.io :as io]
            [failjure.core :as f]
            )
  (:gen-class)) 

;; (def db {:classname  "oracle.jdbc.OracleDriver"
;;          :subprotocol    "oracle:thin"
;;          :subname        "127.0.0.1:1521:orcl" 
;;          :user               "alan"
;;          :password       "alan"})

(defn this-jar
  "utility function to get the name of jar in which this function is invoked"
  [& [ns]]
  (-> (or ns (class *ns*))
      .getProtectionDomain .getCodeSource .getLocation .toURI))

(def db- (:db (read-string (slurp "db.conf"))))

(defn get-db [cclass]
  (let [fname1 (str (-> (java.io.File. (this-jar cclass))
                        .getParentFile .getCanonicalPath)
                    (java.io.File/separator) "db.conf")
        fname2 "db.conf"]
    (if (.exists (io/as-file fname1))
      (:db (read-string (slurp fname1)))
      (:db (read-string (slurp fname2)))
      )))
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
                 B = <'('> A (<','> A)? <')'>
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
                           (insta/transform {:S vector :A str :B vector :AB hash-map})))
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
  ;;(System/exit status)
  )


(def col-names 
  (reduce into (into [] (for [i (cons nil "ABCD")] (into [] (for [j "ABCDEFGHIJKLMNOPQRSTUVWXYZ"] (str i j)))))))

(defn seqb
  ([coll begin] 
   (into [] (drop-while #(not= % (str/upper-case begin)) coll)))
  ([coll begin end] 
   (conj (into [] (take-while #(not= % (str/upper-case end)) (seqb coll begin))) (str/upper-case end))))

(defn sub-rows
  ([coll begin] 
   (drop (dec begin) coll))
  ([coll begin end] 
   (let [cnt (inc (- end begin))] 
     (take cnt (sub-rows coll begin))))) 

(defn row-cell-map [begin-cell-name data]
  (let [[cname rnum] (str/split begin-cell-name #"(?<=[A-Za-z])(?=[0-9])")
        cnames (seqb col-names cname)]
    (map #(map (fn [cn d] [cn %1 d]) cnames %2) (range (Integer. rnum) 100) data)))   

(defn get-sqls [opts]
  (let [sqlfile (:sql opts)
        encoding (det/detect sqlfile "GBK")
        sql (slurp sqlfile :encoding encoding)]
    (and sql (str/split sql #";\s*"))))

(defn get-data [db-spec sqls]
  (reduce conj [] (map (fn [sql] (j/query db-spec [sql] {:as-arrays? true, :keywordize? false}))  sqls)))

(defn update-sheet [sheet start-cell-name data]
  (let [rows (row-cell-map start-cell-name data)]
    (doseq [r rows] 
      (doseq [c r]
        (set-cell! (select-cell (str (nth c 0) (nth c 1)) sheet) (nth c 2))))))

(defn create-wb [wb-name opt data]
  (let [ss-opt (->> opt :sheet)
	paras (reduce into [] (map vector ss-opt data))
        wb (apply create-workbook paras)]
    (save-workbook! wb-name wb)))

(defn update-wb [wb-name opt data]
  (let [wb (load-workbook wb-name)
	ss-opt (->> opt :sheet)
        ss (for [s ss-opt] (if (map? s) (first (vec s)) [s ["A1"]]))]
    (loop [s ss
      	   d data]
      (if-not (first s)
        nil
        (let [sheet (select-sheet (first (first s)) wb)
              first-cell-name (first (second (first s)))
              data (first d)]
          (update-sheet	sheet first-cell-name data)
      	  (recur (rest s) (rest d)))))
    (save-workbook! wb-name wb)))

(defn extract-wb [wb-name opt f]
(let [{:keys [action options exit-message ok?]} 
(validate-args ["-w" "(诊疗项目（门诊+住院_2018）(a1,e10),诊疗项目（仅门诊）)" "-s" "test.sql" "-o" "9087-1.xls" "q2x"])
wb-name (:outfile options)
wb (load-workbook wb-name)
ss-opt (:sheet options)
ss (for [s ss-opt] (if (map? s) (first (vec s)) [s ["A1" "A1"]]))
ins-sqls (read-string (slurp "test-ins.sql"))]
(if (not= (count ss) (count ins-sqls))
 "error: Not equ between sheets and sqls!"
(reduce (fn [ret [s q]] 
  (let [sname (first s)
  		[cell-begin cell-end] (second s)
  		col-begin (re-find #"[A-Za-z]+" cell-begin)
  		col-end (re-find #"[A-Za-z]+" cell-end)
  		row-begin (Integer. (re-find #"[0-9]+" cell-begin))
  		row-end (Integer. (re-find #"[0-9]+" cell-end))
  		col-spec (into {} (map (fn [col name] [(keyword col) (keyword name)])
                        (seqb col-names col-begin col-end) (second ins-sql)))
  		]
        (conj ret (->> wb (select-sheet sname) (select-columns col-spec) (#(sub-rows % row-begin row-end)))) )) [] (map vector ss ins-sql)))))

(defn q2x! [opts db-spec]
  (let [sqls (get-sqls opts)
        sheets (:sheet opts)
        outfile (:outfile opts)]
    (if-not (= (count sqls) (count sheets))
      (f/fail "Error: Not equ between sqls and sheets!")
      (let [data (get-data db-spec sqls)]
        (if (some map? (:sheet opts))
 	  (update-wb outfile opts data)
 	  (create-wb outfile opts data))))))

(defn q2c! [opts db-spec]
	opts)

(defn x2d! [opts db-spec]
	opts)

(defn -main [& args]
  (let [{:keys [action options exit-message ok?]} (validate-args args)
        db (get-db clojure.lang.Atom)]
    (if exit-message
      (exit (if ok? 0 1) exit-message)
      (f/attempt-all
       [ret (case action
              "q2x" (q2x! options db)
              "q2c" (q2c! options db)
              "x2d" (x2d! options db))]
       (println "succeed!")
       (f/when-failed [e]
                      (f/message e))))))
