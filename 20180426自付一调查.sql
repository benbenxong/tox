select * from d_feearea;
select case when out1 >= 50000 and out1 < 100000 then '1_5-10'
            when out1 >= 100000 and out1 < 150000 then '2_10-15'
            when out1 >= 150000 and out1 < 250000 then '3_15-25'
            when out1 >= 250000 and out1 < 350000 then '4_25-35'
            when out1 >= 350000 then '5_35+' end 类别
,count(distinct ic_no) 全市人数
,count(distinct case when feed_area='2260' then ic_no end) 平谷人数
from
(select /*+parallel(a 32) driving_site(a)*/feed_area,ic_no,sum(a.m_in-a.unite_in-a.large_in-supply_in-deformity_in-offi_in) out1
 from combination_stat a
 where st_date >= date'2017-01-01' and out_date >= date'2017-01-01' and out_date < date'2017-12-31'+1
 and m_class not in('11','19')
 group by feed_area,ic_no
) where out1 >= 50000
group by
case when out1 >= 50000 and out1 < 100000 then '1_5-10'
            when out1 >= 100000 and out1 < 150000 then '2_10-15'
            when out1 >= 150000 and out1 < 250000 then '3_15-25'
            when out1 >= 250000 and out1 < 350000 then '4_25-35'
            when out1 >= 350000 then '5_35+' end
order by 1
;

select case when out1 >= 50000 and out1 < 100000 then '1_5-10'
            when out1 >= 100000 and out1 < 150000 then '2_10-15'
            when out1 >= 150000 and out1 < 250000 then '3_15-25'
            when out1 >= 250000 and out1 < 350000 then '4_25-35'
            when out1 >= 350000 then '5_35+' end 类别
,count(distinct ic_no) 全市人数
,count(distinct case when feed_area='2260' then ic_no end) 平谷人数
from
(select /*+parallel(a 32) driving_site(a)*/feed_area,ic_no,sum(a.m_in-a.large_in) out1
 from l_combination_stat a
 where payment_date >= date'2017-01-01' and inh_out_date >= date'2017-01-01' and inh_out_date < date'2017-12-31'+1
 and m_class not in('11','19')
 group by feed_area,ic_no
) where out1 >= 50000
group by
case when out1 >= 50000 and out1 < 100000 then '1_5-10'
            when out1 >= 100000 and out1 < 150000 then '2_10-15'
            when out1 >= 150000 and out1 < 250000 then '3_15-25'
            when out1 >= 250000 and out1 < 350000 then '4_25-35'
            when out1 >= 350000 then '5_35+' end
order by 1
;
