with w as (select first_value(sspweeklyid) over (partition by monthly order by yearweek asc) as id,monthly from sspweekly )
--select distinct * from w where monthly > '2016-02-01' order by monthly asc;
insert into sspmonthlytable(monthly,weekly) (select distinct monthly,id from w where monthly > '2016-02-01' order by monthly asc);
