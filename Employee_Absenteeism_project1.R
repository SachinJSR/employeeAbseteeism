rm(list = ls())
#---Installing the package to read xlx files
#install.packages('xlsx')
library(xlsx)


#Reading the datafile
dataEmpAbs = read.xlsx('d:/edwisor/Absenteeism_at_work_Project.xls',sheetIndex = 1)
dataEmpAbscopy = dataEmpAbs # making a copy of original data
#structure of Data
str(dataEmpAbs)

#variable Names
colnames(dataEmpAbs)

table(dataEmpAbs$Son)
sum(is.na(dataEmpAbs$Month.of.absence))
# splitting the variables into continous and categorical
continuous_vars = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
                   'Work.load.Average.day.', 'Transportation.expense',
                   'Hit.target', 'Weight', 'Height', 
                   'Body.mass.index', 'Absenteeism.time.in.hours')

catagorical_vars = c('ID','Reason.for.absence','Month.of.absence','Day.of.the.week',
                     'Seasons','Disciplinary.failure', 'Education', 'Social.drinker',
                     'Social.smoker', 'Son', 'Pet')

str(dataEmpAbs)
sum(is.na(dataEmpAbs$Son))
unique(dataEmpAbs$Son)
###-----------------------------------------------------------------
# missing value analysis
missVal = data.frame(apply(dataEmpAbs,2,function(x){sum(is.na(x))})) # creating data frame with columns and missing value couunt
missVal$colName = row.names(missVal) # creating a new column with rownames
names(missVal)[1] = 'missCount' # renaming the column 1
row.names(missVal) = NULL # deleting the rownames and new index will be assigned
missVal = data.frame(missVal[,c(2,1)]) # rearranging the columns

library(ggplot2)
ggplot(data = missVal,aes(x = colName,y = missCount)) + geom_bar(stat = 'identity')

#-----------------------------------------
#---Missing value imputation using KNN
library(DMwR)
dataEmpAbs = knnImputation(dataEmpAbs,k=3)

sum(is.na(dataEmpAbs))

str(dataEmpAbs)

#-- Converting categorical variables into Factors
#for (i in catagorical_vars){
#  dataEmpAbs[,i] = factor(dataEmpAbs[,i],labels =(1:length(levels(factor(dataEmpAbs[,i])))))
#}
#-- Converting categorical variables into Factors
for (i in catagorical_vars){
  dataEmpAbs[,i] = as.factor(round(dataEmpAbs[,i]))
}
#-- Converting continous variables into numeric with rounding off
for (i in continuous_vars)
{
  dataEmpAbs[,i] = as.numeric(round(dataEmpAbs[,i]))
}

dataEmpAbsNAfree = dataEmpAbs ## making a copy of data with no null values

table(dataEmpAbs$Month.of.absence)

-------
## Box Plot for OutLier Analysis
  


for (i in 1:length(continuous_vars))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (continuous_vars[i]), x = "Absenteeism.time.in.hours"), data = subset(dataEmpAbs))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=continuous_vars[i],x="Absenteeism.time.in.hours")+
           ggtitle(paste("Box plot of absenteeism for",continuous_vars[i])))
}

# ## Plotting plots together
gridExtra::grid.arrange(gn1,gn2,ncol=2)
gridExtra::grid.arrange(gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,ncol=2)
gridExtra::grid.arrange(gn7,gn8,ncol=2)
gridExtra::grid.arrange(gn9,gn10,ncol=2)

## Replacing the out lier with NA and thn imputing them using KNN

for (i in 1:length(continuous_vars))
{
  val = dataEmpAbs[,i][dataEmpAbs[,i] %in% boxplot.stats(dataEmpAbs[,i])$out]
  dataEmpAbs[,i][dataEmpAbs[,i] %in% val] = NA
}

sum(is.na(dataEmpAbs))

dataEmpAbs = knnImputation(dataEmpAbs,k=3)


##-- feature selection checking corelation
library(corrgram)
library(corrplot)

corrgram(dataEmpAbs,lower.panel = panel.pie,main='Correlation Plot')

corrplot(cor(dataEmpAbs[,continuous_vars]),method = 'color',diag = TRUE,)
# Weight is having redundant data
dataEmpAbs = dataEmpAbs[,!colnames(dataEmpAbs) %in% c('Weight')]  ##Subset can be used too##dataEmpAbs = subset(dataEmpAbs,select = -c('Weight')

## updating the continous Variables
continuous_vars = continuous_vars[-7]

###-----------------------------Feature Scaling
##--- Normality Check

for (i in continuous_vars)
{
  print(i)
  print(var(dataEmpAbs[i]))
}

hist(dataEmpAbs$Absenteeism.time.in.hours)
hist(dataEmpAbs$Distance.from.Residence.to.Work)

#-- Normalization

for (i in continuous_vars)
{
  #print(i)
  dataEmpAbs[i] = (dataEmpAbs[i] - min(dataEmpAbs[i]))/(max(dataEmpAbs[i])-min(dataEmpAbs[i]))
}

## catagorical data -- Creating Dummy variables
#install.packages('dummies')
library(dummies)

dataEmpAbs = dummy.data.frame(dataEmpAbs,catagorical_vars)


###---------------------Model Development
## Data split train and test
library(caTools)
set.seed(101)
sample = sample.split(dataEmpAbs$Absenteeism.time.in.hours,SplitRatio = 0.7)
train = subset(dataEmpAbs,sample == TRUE)
test = subset(dataEmpAbs,sample == FALSE)


####------------------Random Forest
library(randomForest)
## Developing Model with Training Data
rfMOdel = randomForest(Absenteeism.time.in.hours ~., data = dataEmpAbs, importance = TRUE, ntree=100)
# Prediction with test data
rfPredict = predict(rfMOdel,test[,names(test) != 'Absenteeism.time.in.hours'])
# evaluating the Errors
regr.eval(test$Absenteeism.time.in.hours,rfPredict,stats = c('mae','mse','rmse')) #DmWr Library

#        mae         mse        rmse 
#0.020603148 0.003718353 0.060978296  


###----------------------------LinearRegression
#--training model with train data

lrModel = lm(Absenteeism.time.in.hours ~.,train)

summary(lrModel)

lrPredict = predict(lrModel,test[,names(test) != 'Absenteeism.time.in.hours'])

regr.eval(test$Absenteeism.time.in.hours,lrPredict,stats = c('mae','mse','rmse'))
#mae        mse       rmse 
#0.05372474 0.01360402 0.11663626




#------------------------------------------------
ggplot(dataEmpAbscopy,aes(x=Reason.for.absence,y = Absenteeism.time.in.hours))+geom_bar(stat = 'identity')

ggplot(dataEmpAbscopy,aes(x=Month.of.absence,y = Absenteeism.time.in.hours))+geom_bar(stat = 'identity')

ggplot(dataEmpAbscopy,aes(x=Social.drinker,y = Absenteeism.time.in.hours))+geom_bar(stat = 'identity')

ggplot(dataEmpAbscopy,aes(x=Day.of.the.week,y = Absenteeism.time.in.hours))+geom_bar(stat = 'identity')

library(dplyr)
#select(lossDF,Month.of.absence,Absenteeism.time.in.hours,Work.load.Average.day.,Service.time)
lossDF = subset(dataEmpAbsNAfree,select=c('Month.of.absence','Absenteeism.time.in.hours','Work.load.Average.day.','Service.time'))

lossDF$Loss = lossDF$Work.load.Average.day.*lossDF$Absenteeism.time.in.hours / lossDF$Service.time



month = group_by(lossDF,lossDF$Month.of.absence)
result = summarise(month,lossSum = sum(Loss))
result

data.frame(row.names = c('zero','jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'),Loss = result[2])



