setwd("~/Predictive Analytics/lending matter pricing")
library(xlsx)
responses = read.xlsx("responses_N53.xlsx", sheetIndex = 1)
timecard = read.xlsx("TimeCardData_lendingmatters_N53.xlsx", sheetIndex = 1)

data = subset(timecard, timecard$section=="Equity Partner" | timecard$section=="Non-Equity Partner" | timecard$section=="Principal" | timecard$section=="Partner" | timecard$section=="Contract Partner" | timecard$section=="Counsel" | timecard$section=="Other Counsel")

keeps <- c("Number", "PPLcount", "sum_hours", "sum_amt")
numerical_data = data[ , (names(data) %in% keeps)]
str(numerical_data)
agg_data = aggregate(. ~ Number, numerical_data, sum)

idx = which(names(responses)=="Please.enter.the.matter.number")
names(responses)[idx] = "Number"

data_merged = merge(agg_data, responses, by="Number", all=T)

# write.xlsx(data_merged, file="data_merged_partners.xlsx")
data_merged = data_merged[!duplicated(data_merged$Number),]
drop1 = which(names(data_merged)=="Item.Type")
drop2 = which(names(data_merged)=="Path")
drop3 = which(names(data_merged)=="When.did.the.transation.close.")
drop4 = which(names(data_merged)=="Created.By")
drop5 = which(names(data_merged)=="sum_amt")
drop6 = which(names(data_merged)=="Did.the.transaction.close.")
drop7 = which(names(data_merged)=="Number")
data_merged = data_merged[,-c(drop1,drop2,drop3,drop4,drop5,drop6,drop7)]
sec_doc = which(names(data_merged)=="Please.describe.the.nature.of.the.security.document")
lead = which(names(data_merged)=="For.the.cross.border.international.transaction..was.McMillan.the.lead.counsel.or.local.support.counsel.")

NAcount = sapply(data_merged, function(x) sum(is.na(x)))
complete_cols = which(NAcount==0)
data_merged = data_merged[,c(complete_cols,sec_doc, lead)]
NA_dollar = which(data_merged$What.was.the.dollar.amount.of.the.loan.in.millions.=="NA")
data_merged = data_merged[-NA_dollar,]
data_merged$What.was.the.dollar.amount.of.the.loan.in.millions.=as.numeric(data_merged$What.was.the.dollar.amount.of.the.loan.in.millions.)
NAcount = sapply(data_merged, function(x) sum(is.na(x)))
NAcols = which(NAcount>0)
data_merged[,NAcols[1]] = as.character(data_merged[,NAcols[1]])
data_merged[,NAcols[2]] = as.character(data_merged[,NAcols[2]])
data_merged[,NAcols[1]] = replace(data_merged[,NAcols[1]],which(is.na(data_merged[,NAcols[1]])==T),"N/A")
data_merged[,NAcols[2]] = replace(data_merged[,NAcols[2]],which(is.na(data_merged[,NAcols[2]])==T),"N/A")
data_merged[,NAcols[1]] = as.factor(data_merged[,NAcols[1]])
data_merged[,NAcols[2]] = as.factor(data_merged[,NAcols[2]])
str(data_merged)


df1_dummy <- dummyVars("~.",data=data_merged, fullRank=F)
df1_dummy <- data.frame(predict(df1_dummy, newdata = data_merged))

lmfit = lm(sum_hours~., df1_dummy)
summary(lmfit)

# write.xlsx(data_merged,file="data_complete.xlsx", row.names = F)
nature_of_security_doc = data_merged$Please.describe.the.nature.of.the.security.document
N=length(nature_of_security_doc)
security_type = data.frame(matrix(0, nrow=N, ncol = 12))
names(security_type) = c("Guarantee","General Security Agreement","Share/note pledge Agreement","Hypothec - notarized",
                         "Hypothec - not notarized", "Mortgage", "Control", "IP secuirty", "Suborfination", "Bank act", "Other", "N/A")
security_type[,1] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Guarantee",x)))),0,1) 
security_type[,2] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("General Security Agreement",x)))),0,1)
security_type[,3] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Share/note pledge Agreement",x)))),0,1)
security_type[,4] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Hypothec - notarized",x)))),0,1)
security_type[,5] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Hypothec - not notarized",x)))),0,1)
security_type[,6] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Mortgage",x)))),0,1)
security_type[,7] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Control",x)))),0,1)
security_type[,8] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("IP security",x)))),0,1)
security_type[,9] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Subordination",x)))),0,1)
security_type[,10] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Bank act",x)))),0,1)
security_type[,11] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("Other",x)))),0,1)
security_type[,12] = ifelse(is.na(as.numeric(sapply(nature_of_security_doc, function(x) grep("N/A",x)))),0,1)

drops <- "Please.describe.the.nature.of.the.security.document"
data_merged = data_merged[ , ! (names(data_merged) %in% drops)]
data_merged = cbind(data_merged, security_type)
str(data_merged)

nzv = nearZeroVar(data_merged)
print(names(data_merged)[nzv])

drop = which(names(data_merged)=="In.what.industry.does.the.borrower.carry.on.business.")
data_merged = data_merged[,-drop]
names(data_merged)
# include = c("PPLcount","sum_hours", "Was.the.loan.secured.", "Was.this.an.international.or.cross.border.transaction.",
#             "Who.performed.the.corporate.and.PPSA.RPMRR.searches.and.registrations.",
#             "What.was.the.dollar.amount.of.the.loan.in.millions.",
#             "What.type.of.loan.agreement.was.provided.",
#             "Please.describe.the.type.of.the.lending.transaction",
#             "Was.this.matter.part.of.an.acquisition.or.other.corporate.transaction.that.we.were.involved.in.",
#             "For.the.cross.border.international.transaction..was.McMillan.the.lead.counsel.or.local.support.counsel.")
# data_merged = data_merged[ , (names(data_merged) %in% include)]


#############################################################
# Data partitioning
############################################################
set.seed(1)
inTrain = createDataPartition(data_merged$sum_hours, p=0.8, list=F)
data_tr = data_merged[inTrain,]
data_te = data_merged[-inTrain,]

#############################################################
# Linear model
############################################################

lmfit = lm(sum_hours~., data=data_tr)
summary(lmfit)
lmfit = step(lmfit, direction="both")
summary(lmfit)
coefs = summary(lmfit)
write.xlsx(coefs$coefficients, file="coefs.xlsx")
pred = predict(lmfit, newdata=data_te)
actual = data_te$sum_hours
MAPE = 100*mean(abs(pred-actual)/actual)
print(MAPE)

##############################################################
# random forest
#############################################################
rf_control <- trainControl(method="repeatedcv", number=5, repeats=1)
seed <- 7
set.seed(seed)
M = dim(data_merged)[2]-1

rf = train(x=data_tr[,-2],
           y=data_tr[,2], 
           data=data_merged, 
           method="rf", 
           tuneLength = M)
plot(rf)
summary(rf)
print(rf)

pred = predict(rf, newdata=data_te)
actual = data_te$sum_hours
MAPE = 100*mean(abs(pred-actual)/actual)
print(MAPE)

pred = predict(rf, newdata=data_tr)
actual = data_tr$sum_hours
MAPE = 100*mean(abs(pred-actual)/actual)
print(MAPE)

#########################################
# gbm 
#######################################
gbmgrid = expand.grid(n.trees = c(100,500,1000),
                      shrinkage = c(0.001, 0.01, 0.1),
                      interaction.depth = c(1,2,4,6,8,10),
                      n.minobsinnode = 10)

gbmControl = trainControl(method="cv", number = 5, verboseIter = F, allowParallel = T)

gbm = train(x=as.matrix(data_tr[,-2]),
            y=data_tr[,2],
            method="gbm",
            tuneGrid = gbmgrid)
                      
summary(gbm)
print(gbm)


pred = predict(gbm, newdata=data_te)
actual = data_te$sum_hours
MAPE = 100*mean(abs(pred-actual)/actual)
print(MAPE)

pred = predict(gbm, newdata=data_tr)
actual = data_tr$sum_hours
MAPE = 100*mean(abs(pred-actual)/actual)
print(MAPE)


# 