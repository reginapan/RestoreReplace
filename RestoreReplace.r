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