SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  
sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN (SELECT custid, agent,
max(tgl) as tglmax FROM mgm_hst 
where tgl BETWEEN  '2008-06-23 0:00:00' and '2008-06-23 23:59:00'  
and  custid in(select custid from mgm where agent in 
(select userid from usertbl where  SPVCODE >= 'SPV1' and SPVCODE <= 'SPV1')) 
group by custid,agent) as a  on mgm_hst.custid = a.custid 
and mgm_hst.tgl = a.tglmax  inner join mgm on mgm.custid = a.custid 
and a.agent=mgm.agent where  recsource BETWEEN '-----' and 'ZZZZZ' 
AND MGM_HST.F_CEK IN ('NK-HO','NK-OF','MV-EC','MV-HO','MV-OF','MV-EC',
'MV-OV','WN-HO','WN-OF','WN-EC','NA-Z','NA-G-','NA-H','NA-P',
'NA-M','NA-O','RP-B','RP-E','RP-F','RP-R','RP-J',
'RP-L','RP-M','RP-N','RP-Q','RP-T','RP-U','RP-W','BP-','OP-','PTP','POP','ST-','NBP')
group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent 

SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  
sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN (SELECT custid, agent,
max(tgl) as tglmax FROM mgm_hst 
where tgl BETWEEN  '2008-06-23 0:00:00' and '2008-06-23 23:59:00'  
and  custid in(select custid from mgm where agent in 
(select userid from usertbl where  SPVCODE >= 'SPV1' and SPVCODE <= 'SPV1')) 
group by custid,agent) as a  on mgm_hst.custid = a.custid 
and mgm_hst.tgl = a.tglmax  inner join mgm on mgm.custid = a.custid 
and a.agent=mgm.agent where  recsource BETWEEN '-----' and 'ZZZZZ' 
AND (left(MGM_HST.F_CEK,2) IN ('NK','MV','WN','NA',
'RP','BP','OP','ST') or left(MGM_HST.F_CEK,3) in ('NBP','PTP','POP')) 
group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent 


SELECT F_CEK from MGM where agent='devel1'