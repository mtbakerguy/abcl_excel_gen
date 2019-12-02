((:name "a"
	:rows '((:row-values '("name" "age"))
		(:row-values '("mickey" 50))
		(:row-values '("minnie" 48 (:color 'green :cell-value "~a+~a" :formula '((-1 -1) (-1 0)))))
                (:row-values  '((:custom 'goofy)))
                (:row-values  '((:custom 'dopey)))
		(:row-values '("goofy" 48 (:bold t :strikeout t :italic t :color 'red :cell-value "bogus")))
		(:row-values '("goofy" 48 (:bold t :strikeout t :italic t :color 'orange :cell-value "bogus")))
		(:row-values '("goofy" 48 (:bold t :strikeout t :italic t :color 'yellow :cell-value "bogus")))
		(:row-values '("goofy" 48 (:bold t :strikeout t :italic t :color 'green :cell-value "bogus")))
		(:row-values '("goofy" 48 (:bold t :strikeout t :italic t :color 'blue :cell-value "bogus")))
		(:row-values '("goofy" 48 (:bold t :strikeout t :italic t :color 'indigo :cell-value "bogus")))
		(:row-values '("goofy" 48 (:bold t :strikeout t :italic t :color 'violet :cell-value "bogus")))
		(:row-values '("mickey jr" 8 nil (:bold t :cell-value 44 :bottom t :top t :right t :left t)))))
 (:name "b"
	:rows '((:row-values '("name" "age"))
                (:hidden t :row-values '(nil 0))
		(:row-values '("scrooge" 50))
		(:row-values '("huey" 10))
		(:row-values '("dewey" 10))
		(:row-values '("louie" 10)))))


	
