import xlrd

#Venue File Opening
file_location_venue='Venue.xlsx'
workbook_1=xlrd.open_workbook(file_location_venue)
sheet_1=workbook_1.sheet_by_index(0)

#Match File Opening
file_location_match='Match.xlsx'
workbook_2=xlrd.open_workbook(file_location_match)
sheet_2=workbook_2.sheet_by_index(0)

#input season
season_no = int(input("Enter the season number:"))
print('\n')

x=sheet_2.nrows-1   
y=sheet_1.nrows-1

team_codes=[]
team_playing_code=[]
home_flag=0
match_no=1
away_count=0
home_count=0
home_list=[]
home_toss=0
away_toss=0
toss_team=0
host_teamwon_flag=0
host_teamwon=0
bat_teamwon=0
tosslost=0

for a in range(x):
	if(sheet_2.cell_value((a+1),4)==season_no):
		print '--MATCH NO:',match_no,'--'
		for b in range(y):
			if(sheet_2.cell_value((a+1),5)==sheet_1.cell_value((b+1),0)):

				team_codes=[int(sheet_1.cell_value((b+1),1)),int(sheet_1.cell_value((b+1),2)),int(sheet_1.cell_value((b+1),3))]
				team_playing_code=[int(sheet_2.cell_value((a+1),2)),int(sheet_2.cell_value((a+1),3))]
				print(team_codes)
				print(team_playing_code)
				for c in team_codes:
					for d in team_playing_code:
						if(c==d):
							print 'HOME For: '												
							home_list.append(d)							
							home_flag=1


				
				if(home_flag == 0):
					print 'AWAY !'
					away_count=away_count+1	

				print(home_list)
				

				for e in home_list:
				 	if(e==sheet_2.cell_value((a+1),6)):
				 		print 'toss won by home:',e
						home_toss=home_toss+1
					else:
						print'toss won by away:'
						away_toss=away_toss+1

				#toss won +team won		
				if (sheet_2.cell_value((a+1),3) == sheet_2.cell_value((a+1),13)):
					toss_team=toss_team+1

				#home team + team won
				for f in home_list:
					if(f == sheet_2.cell_value((a+1),13)):
						host_teamwon_flag=1


				if ( host_teamwon_flag == 1):
					host_teamwon=host_teamwon+1


				if (sheet_2.cell_value((a+1),7) == 'bat' and (sheet_2.cell_value((a+1),3) == sheet_2.cell_value((a+1),13))):
					bat_teamwon=bat_teamwon+1

				# if(sheet_2.cell_value((a+1),2)!=sheet2.cell_value((a+1),6) and sheet2.cell_value((a+1),3)==sheet2.cell_value((a+1),13))
				# 	tosslost=tosslost+1


				host_teamwon_flag=0
				home_list=[]	
 				home_flag=0
				print("\n")
				team_codes=[]
				team_playing_code=[]

		match_no=match_no+1


print '\n\n'
print'SEASON',season_no,'SUMMARY'

print'total matches:',match_no
home_count=match_no - away_count
print'total home matches played on either of the teams pitch:',home_count
print'total away matches,neither of their pitches:',away_count
#print'matches played on neither of home or away ground',(match_no-(home_toss+away_toss))
print'toss won + home teams:',home_toss
print'toss won + team won: ',toss_team
print 'host team + team won: ',host_teamwon
print 'toss won +team won + bat first team: ',bat_teamwon
#print'toss won by away teams:',away_toss
print 'toss lost by team 1, match won ',tosslost
