########################################################################
#  rm(list=ls())��տ���̨���б���                                     #
#  .libPaths()�鿴�����·��                                           #
#  openxlsx��Ҫ����Rtools,����:                                        #
#  1����װinstallr����install.packages("installr")��                   #
#  2��ִ������installr::install.rtools()                               #
#  3��������ɺ�װ,��װ�����г��˻������������������ȫ��Ĭ�ϼ���  #
########################################################################

#���ù���Ŀ¼
setwd("E:/QD/tools/Linear-Solver/V2")
library(Rglpk)
library(openxlsx)
#��ȡ�����Լ�����
all<-read.csv("raw_data.csv",header=T)
raw_data<-all[,2:ncol(all)]
quota_limit<-read.csv("quota_limit.csv",header=T)
quota_limit$condition<-as.character(quota_limit$condition)
#�����б�������
rd_col<-ncol(raw_data)
#������������ޣ�Ĭ�����¸���1%
float<-round(sum(quota_limit$quota)/rd_col*0.01)
quota_limit$max<-quota_limit$quota+float
quota_limit$min<-quota_limit$quota-float

#����ͳ�Ʋ����ɱ���x
data<-aggregate(raw_data[,1],raw_data[,1:rd_col],length)
#������������������
names(data)[rd_col+1]<-"raw"
#���������б���x��ֵΪx1...xi
for(i in 1:length(data$raw)){
  data$x[i]<-paste("x",i,sep="")
}
#������������
n<-nrow(quota_limit)
#ÿ��������Ӧ��xi����
for(i in 1:n){
  assign(paste("con",i,sep=""),eval(parse(text=paste(quota_limit[i,1]))))
}
#����m*n�ľ���m��������������������n�����ɱ䵥Ԫ��������
aa<-(paste("x",seq(1:nrow(data)),sep=""))
bb<-matrix(c(rep(aa,n)),nrow=n,byrow=T)

#�Խ�����
rsl<-function(data){
  obj<-c(rep(1,nrow(data)))
  for(i in 1:n){
    bb[i,which(bb[i,] %in% eval(parse(text=paste("con",i,sep=""))))]<-1 #��coni���е�Ԫ�������滻Ϊ1
    bb[i,which(regexpr("x",bb[i,])>0)]<-0 #������"x"�ĵ�Ԫ�������滻Ϊ0
  }
  bb1<-bb
  bb2<-rbind(bb,bb1)
  cc<-diag(1,nrow=nrow(data))#�ɱ䵥Ԫ���ϵ������
  mat<-rbind(bb2,cc)#������ϵ������
  dir<-c(rep(c("<=",">="),each=n),rep("<=",nrow(data)))#����
  rhs1<-append(quota_limit$max,quota_limit$min)
  rhs<-t(append(rhs1,data$raw))#ת�������������(append:׷�����ݣ����ﲻ����rbind)
  ls<-Rglpk_solve_LP(obj,mat,dir,rhs,max=T,types = "I")#���
  data$combine<-do.call(paste,c(data[,1:(ncol(data)-2)],sep=""))#��ϱ���
  result<-cbind(data,ls$solution)
  names(result)[ncol(result)]<-"solution"
  return(result)
}
result<-rsl(data)#�滮���

#����Թ滮�����ѡ����
combine<-do.call(paste,c(all[,2:ncol(all)],sep=""))
#�Զ��庯��������������ֻͳ�Ƶ�Ŀǰ��¼���еģ�����ȫ��
cf<-function(array){
  accum<- c()
  n<-rep(1,length(array))#����ֻ���ۼ����㣬������ȫ��Ϊ1��n���������ۼ�
  temp<-tapply(n,array,cumsum)#��array�ۼ�
  a<-sort(array[!duplicated(array)])#ȥ�ز�˳������
  i<-length(a)
  for(i in 1:i){
    accum[which(array==a[i])] <- temp[[i]]#���ۼ�ֵ��������
  }
  accum<-cbind(array,accum)
  return(as.data.frame(accum))
}
accum<-cf(combine)
names(accum)[1]<-"combine"
all<-cbind(all,accum)
all<-merge(all,result[,(ncol(result)-1):(ncol(result))],by="combine")
#������С�ڹ滮������Ϊ1������Ϊ0
all$select[as.numeric(as.character(all$accum))<=all$solution]<-1#���ӱ��뾭��˫��ת��Ϊ��ֵ
all$select[as.numeric(as.character(all$accum))>all$solution]<-0
#Ŀ��ֵ�͹滮ֵ�Ա�
myfun<-function(df,default){
  if(default==1){
    nc<-ncol(result)
    df1<-gsub("\\$x","",df) #"//":ת���,����$������Ϊ�������ʽ
    df2<-gsub("data","result",df1)
    df3<-gsub("]",paste(",",nc,"]",sep=""),df2)
  }
  if(default==0){
    nc<-ncol(result)-3
    df1<-gsub("\\$x","",df) #"//":ת���,����$������Ϊ�������ʽ
    df2<-gsub("data","result",df1)
    df3<-gsub("]",paste(",",nc,"]",sep=""),df2)
  }
  return(df3)
}
compare<-quota_limit[,1:2]
compare$raw<-0
compare$raw_gap<-0
compare$plan<-0
compare$plan_gap<-0

for(i in 1:n){
  compare[i,3]<-sum(eval(parse(text=myfun(compare[i,1],0))))
  compare[i,4]<-compare[i,3]-compare[i,2]
  compare[i,5]<-sum(eval(parse(text=myfun(compare[i,1],1))))
  compare[i,6]<-compare[i,5]-compare[i,2]
  compare[i,1]<-gsub("data\\$x","",compare[i,1])
}

#ȥ���м����
samples_select<-subset(all,select=-c(combine,accum,solution))
result<-subset(result,select=-c(x,combine))
#������
wb<-createWorkbook("Creator of workbook")#����һ���¹�����
sheet_name<-c("samples_select","result","compare")

for (i in 1:3) {
  addWorksheet(wb, sheetName = sheet_name[i])#�ڴ����Ĺ�����������sheet
  writeData(wb,sheet=i,eval(parse(text=sheet_name[i])))#�ֱ�д�����ݵ�sheet��
  freezePane(wb, sheet_name[i] ,  firstActiveRow =2)#��������
  if(i!=3){
    addStyle(wb,sheet=i, createStyle(fgFill="sienna2",halign="center"), rows=1, cols=1:ncol(eval(parse(text=sheet_name[i]))))
    addStyle(wb,sheet=i, createStyle(halign="center"), rows=2:(nrow(eval(parse(text=sheet_name[i])))+1), cols=1:6,gridExpand=T)
    
  }else{
    addStyle(wb,sheet=i, createStyle(fgFill="sienna2",halign="center"), rows=1, cols=1:6)
    addStyle(wb,sheet=i, createStyle(halign="center"), rows=2:(nrow(eval(parse(text=sheet_name[i])))+1), cols=1:6,gridExpand=T)
    setColWidths(wb, sheet=i, cols=1, widths = "auto")
  }
}

saveWorkbook(wb, file = "final.xlsx", overwrite = TRUE) #���excel�ļ�