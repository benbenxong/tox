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

(def db (:db (read-string 
              (slurp (str (System/getenv "temp")
                          (java.io.File/separator)
                          "tox" (java.io.File/separator)
                          "db.conf")))))

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
	"  d2x    db to excel file."
	"  d2t    db to txt|csv file(s)."
	"  x2d    excel file to db."
	""
	"Please refer to the manual page for more information."]
       (str/join \newline)))

(defn error-msg [errors]
  (str "The following errors occurred while parsing your command:\n\n"
       (str/join \newline errors)))

(def actions #{"d2x" "x2d" "d2t" "x2s"})

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
      (some actions arguments)
      {:action (some actions arguments) :options options :arguments arguments}
      :else ; failed custom validation => exit with usage summary
      {:exit-message (usage summary)})))

(defn exit [status msg]
  (println msg)
  (System/exit status)
  )

(defn mapc [f coll] (reduce (fn [_ v] (f v)) nil coll))

(def col-names 
  (reduce into (into [] (for [i (cons nil "ABCD")] (into [] (for [j "ABCDEFGHIJKLMNOPQRSTUVWXYZ"] (str i j)))))))

(defn seqb
  ([coll begin] 
   (into [] (drop-while #(not= % (str/upper-case begin)) coll)))
  ([coll begin end] 
   (if (nil? end) 
     (seqb coll begin)
     (conj (into [] (take-while #(not= % (str/upper-case end)) (seqb coll begin))) (str/upper-case end)))))

(defn sub-rows
  ([coll begin] 
   (drop (dec begin) coll))
  ([coll begin end] 
   (if (nil? end) 
     (sub-rows coll begin)
     (let [cnt (inc (- end begin))] 
       (take cnt (sub-rows coll begin)))))) 

(defn row-cell-map [begin-cell-name data]
  (let [[cname rnum] (str/split begin-cell-name #"(?<=[A-Za-z])(?=[0-9])")
        cnames (seqb col-names cname)]
    (map #(map (fn [cn d] [cn %1 d]) cnames %2) (range (Integer. rnum) 100) data)))   

(defn get-sqls [sql-file args]
  (let [encoding (det/detect sql-file "GBK")
        sql (slurp sql-file :encoding encoding)
        sql-args (into #{} (re-seq #"&\d" sql))
        cnt-sql-args (count sql-args)
        cmd-args (into [] (filter (fn [x] (->> x actions not)) args))
        cnt-args (count cmd-args)
        replace-map (into {} (map (fn [i v] [(str "&" i) v]) (range 1 10) cmd-args))]
    (if (not= cnt-args cnt-sql-args)
      (f/fail (str "Not equ between cnt-args[" cnt-args "] and cnt-sql-args[" cnt-sql-args "]!"))
      (if (= 0 cnt-args) 
        (and sql (str/split sql #";\s*"))
        (and sql (str/split (str/replace sql #"&\d" replace-map) #";\s*"))))))

;; (defn get-data [db-spec sqls]
;;   (reduce conj [] (map (fn [sql] (j/query db-spec [sql] {:as-arrays? true, :keywordize? false}))  sqls)))

(defn get-ins-sqls [ins-sql-file]
  (read-string (slurp ins-sql-file)))

;; (->> (get-sqls "test.sql") (db-to-seq db))
(defn db-to-seq [db-spec sqls]
  (into []
        (map-indexed 
         (fn [ind sql]
           (println sql) 
           [ind (j/query db-spec [sql] {:as-arrays? true, :keywordize? false})]) sqls)))

;; (let [{:keys [action options exit-message ok?]} (validate-args [
;; "-o" "test-temp2.xlsx" "-w" "(sheet1(a2,b3),sheet2(a2,b3))" "-s" "test-ins.sql" "d2x"])]
;; (xls-to-seq (:outfile options) (:sheet options) (get-ins-sqls "test-ins.sql"))
;; ;;(println (:outfile options) "," (:sheet options))
;; ;;(map vector (reduce into [] (:sheet options)) (get-ins-sqls "test-ins.sql"))
;; )
(defn xls-to-seq [wb-name wb-sheets ins-sqls]
  (let [wb (load-workbook wb-name)
  	ss (reduce into [] wb-sheets)]
    (if-not (every? map? wb-sheets)
      (f/fail "error: When xls-to-seq sheet Must be Map type (sheet(a1,a2)..)!")
      (into [] (map-indexed 
                vector 
                (reduce (fn [ret [s q]] 
                          (let [sname (first s)
  		                [cell-begin cell-end] (second s)
  		                col-begin (re-find #"[A-Za-z]+" cell-begin)
  		                col-end (if cell-end (re-find #"[A-Za-z]+" cell-end))
  		                row-begin (Integer. (re-find #"[0-9]+" cell-begin))
  		                row-end (if cell-end (Integer. (re-find #"[0-9]+" cell-end)))
  		                col-spec (into {} (map (fn [col name] [(keyword col) (keyword name)])
                                                       (seqb col-names col-begin col-end) (second q)))]
                            (conj ret
                                  (into [] (cons (mapv name (second q)) (->> wb 
                                                                             (select-sheet sname) 
                                                                             (select-columns col-spec) 
                                                                             (#(sub-rows % row-begin row-end))
                                                                             (map (fn [r] (vec (vals r))))))))))
                        []
                        (map vector ss ins-sqls)))))))

;; (let [{:keys [action options exit-message ok?]} (validate-args [
;;  "-o" "test-temp2.xlsx" "-w" "(sheet1,sheet2)" "-s" "test-ins.sql" "d2x"])
;;   ss (:sheet options)
;;   data  (db-to-seq db (get-sqls "test.sql"))]
;;  (seq-to-xls (:outfile options) (:sheet options) data)
;;  ) 
(defn seq-to-xls [wb-name wb-sheets seq-data] 
  (if (not= (count wb-sheets) (count seq-data))
    (f/fail "Not equ count between wb-sheets and seq-data!")
    (if (some map? wb-sheets)
      ;;update-workbook
      (let [wb (load-workbook wb-name)
	    ss (reduce into [] wb-sheets)]
	(if (not= (count ss) (count seq-data))
	  (f/fail "Not equ count between wb-sheets and seq-data")
    	  (do
    	    (doseq [[s d] (map vector ss seq-data)]
              (let [sheet (select-sheet (first s) wb)
                    first-cell-name (first (second s))
                    col-begin (re-find #"[A-Za-z]+" first-cell-name)
                    row-begin (Integer. (re-find #"[0-9]+" first-cell-name))
                    data (->> d second rest)]
                (doseq [[row-data row-num] (map vector data (range row-begin 1000000))]
          	  (doseq [[col-name cell-data] 
          	          (map vector (seqb col-names col-begin)
          	               row-data)]
          	    (set-cell! (select-cell (str col-name row-num) sheet) cell-data)))))
            (save-workbook! wb-name wb))))
      ;;create-workbook
      (let [paras (reduce into [] (map (fn [s d] [s (second d)]) wb-sheets seq-data))
            wb (apply create-workbook paras)]
        (save-workbook! wb-name wb)
        ))))

(defn seq-to-db [db-spec ins-sqls seq-data]
  (mapc (fn [[sql sdata]] 
        (j/insert-multi! 
         db-spec 
         (first sql) 
         (->> sdata second first) 
         (->> sdata second rest)) sql) (map vector ins-sqls seq-data)))

(defn- update-sheet [sheet start-cell-name data]
  (let [rows (row-cell-map start-cell-name data)]
    (doseq [r rows] 
      (doseq [c r]
        (set-cell! (select-cell (str (nth c 0) (nth c 1)) sheet) (nth c 2))))))

(defn- create-wb [wb-name wb-sheets data]
  (let [paras (reduce into [] (map vector wb-sheets data))
        wb (apply create-workbook paras)]
    (save-workbook! wb-name wb)))

(defn- update-wb [wb-name opt data]
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
          (update-sheet	sheet first-cell-name (first data))
      	  (recur (rest s) (rest d)))))
    (save-workbook! wb-name wb)))

(defn- extract-wb [opt db-spec f]
  (let [wb-name (:outfile opt)
        wb (load-workbook wb-name)
        ss-opt (:sheet opt)
        ss (for [s ss-opt] (if (map? s) (first (vec s)) [s ["A1" "A1"]]))
        sql-file (:sql opt)
        ins-sqls (read-string (slurp sql-file))]
    (if (not= (count ss) (count ins-sqls))
      (f/fail "error: Not equ between sheets and sqls!")
      (let [data 
            (reduce (fn [ret [s q]] 
                      (let [sname (first s)
  		            [cell-begin cell-end] (second s)
  		            col-begin (re-find #"[A-Za-z]+" cell-begin)
  		            col-end (if cell-end (re-find #"[A-Za-z]+" cell-end))
  		            row-begin (Integer. (re-find #"[0-9]+" cell-begin))
  		            row-end (if cell-end (Integer. (re-find #"[0-9]+" cell-end)))
  		            col-spec (into {} (map (fn [col name] [(keyword col) (keyword name)])
                                                   (seqb col-names col-begin col-end) (second q)))]
                        (conj ret
                              (->> wb 
                                   (select-sheet sname) 
                                   (select-columns col-spec) 
                                   (#(sub-rows % row-begin row-end))))))
                    [] 
                    (map vector ss ins-sqls))]
        (f db-spec ss ins-sqls data)))))

(def data-to-db 
  (fn [db-spec _ ins-sqls data]
    (mapc (fn [[sql sdata]] (j/insert-multi! db-spec (first sql) sdata) nil) (map vector ins-sqls data))))

(def data-to-seq
  (fn [_ ss _ data]
    (mapc (fn [[sheet sdata]] {(first sheet) sdata}) (map vector ss data))))

(defn d2x! [opts db-spec args]
  (let [sql-file (:sql opts)
        wb-sheets (:sheet opts)
        wb-name (:outfile opts)]
    (if-not (and sql-file wb-sheets wb-name)
      (f/fail "Paras error: d2x Needs :sql :sheet :outfile!")
      (let [sqls (get-sqls sql-file args)]
            (->> (db-to-seq db-spec sqls) (seq-to-xls wb-name wb-sheets))))))

(defn x2d! [opts db-spec]
  (let [sql-file (:sql opts)
        wb-sheets (:sheet opts)
        wb-name (:outfile opts)]
    (if-not (and sql-file wb-sheets wb-name)
      (f/fail "Paras error: x2d Needs :sql :sheet :outfile!")
      (let [ins-sqls (get-ins-sqls sql-file)]
        (->> (xls-to-seq wb-name wb-sheets ins-sqls) (seq-to-db db-spec ins-sqls))))))

(defn db-to-txt [db-spec sqls file-name]
  (->> (db-to-seq db-spec sqls) (#(str/replace % #"]" "]\r\n")) (spit file-name)))

(defn d2t! [opts db-spec args]
  (let [sql-file (:sql opts)
        outfile (:outfile opts)]
    (if-not (and sql-file outfile) 
      (f/fail "Paras error: d2t Needs :sql :outfile!")
      (let [sqls (get-sqls sql-file args)]
        (db-to-txt db-spec sqls outfile)))))

(defn x2s! [opts]
  (let [sql-file (:sql opts)
        ins-sqls (get-ins-sqls sql-file)
        wb-sheets (:sheet opts)
        wb-name (:outfile opts)]
    (if-not (and sql-file wb-sheets wb-name)
      (f/fail "Paras error: x2s Needs :sql :sheet :outfile!")
      (let [ins-sqls (get-ins-sqls sql-file)]
        (xls-to-seq wb-name wb-sheets ins-sqls )))))

(defn -main [& args]
  (let [{:keys [action options arguments exit-message ok?]} (validate-args args)]
    (if exit-message
      (exit (if ok? 0 1) exit-message)
      (f/attempt-all
       [ret (case action
              "d2x" (d2x! options db arguments)
              "x2d" (x2d! options db)
              "x2s" (x2s! options)
              "d2t" (d2t! options db arguments))]
       (println "succeed!")
       (f/when-failed [e]
                      (exit 1 (f/message e)))))))
