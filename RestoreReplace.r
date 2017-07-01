library(xlsx)
ref<-read.xlsx(file='reference.xlsx',sheetIndex = 1)
sch<-read.xlsx(file='schedule.xlsx',sheetIndex = 1)

extract<-function(x)
{data<-x[c(1:3)]
data$date<-paste("2017/7/",1,sep="")
names(data)[3]<-c("status")
for(i in 4:33){
datatemp<-x[c(1,2,i)]
datatemp$date<-paste("2017/7/",i,sep="")
names(datatemp)[3]<-c("status")
data<-rbind(data,datatemp)
}
return(data)
}
data<-extract(sch)
isNA<-subset(data,is.na(data[c(3)] == TRUE))
isduty<-subset(data,is.na(data[c(3)] == FALSE))
write.xlsx(isNA,"NAs.xlsx")
write.xlsx(isduty,"isduty.xlsx")

ref$empid<-as.character(ref$empid)
empidlist<-as.character(isNA$empid)

findreplace<-function(x){
	ref$empid<-as.character(ref$empid)
	replaceid<-c(NA)
	num<-c(0)
	replaceset<-data.frame(replaceid,num)
	for(i in 1:158){
		id<-x[i]
		temp<-subset(ref,(ref$empid == id)==TRUE,select = c(replaceid,ratio))
		if((is.data.frame(temp) && nrow(temp)==0)==FALSE)
		{replaceid<-subset(temp,temp$ratio == max(ratio),select = replaceid)
		replaceid<-data.frame(replaceid)
		replaceid$num<-i
		replaceset<-rbind(replaceset,replaceid)
		}
		else{
		replaceid<-NA
		replaceid<-data.frame(replaceid)
		replaceid<-i
		replaceset<-rbind(replaceset,replaceid)
		}
	}
	return(replaceset)
}

Listofreplace<-findreplace(empidlist)
write.xlsx(Listofreplace,"Result")
