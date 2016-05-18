######################################################
# Permeability Data Processing                       #
#   - remove non informative rows                    #                       
#   - select standards i.e. rows starting with "Mix" #
#   - select specific columns                        #
#   - plots                                          #
#                                                    #
# Diana Coman Schmid                                 #
# diana.comanschmid@eawag.ch                         #
# 10.04.2015                                         #
######################################################

#install.packages("rJava")
#install.packages("XLConnect")
library("XLConnect")
library("Hmisc")

# specify constants
# can be flexibly changed

mw.rho <- 1000
mw.ly <- 444.25

C0.ab.rho <- 700
C0.ab.ly <- 50
V.ab.adon <- 3.5
V.ab.brec <- 4

C0.ba.rho <- 612
C0.ba.ly <- 40
V.ba.arec <- 3.5
V.ba.bdon <- 4

A <- 4.52
V <- 0.15

# specify the data location in the excel sheets:
    # rows with data: for "ins" inserts => "30:30+m+1" and "60:60+m+1" (machine exports data per default starting 
    # with row 30 for the first dye and row 60 for the second dye)
    # columns with data: for "rep" technical replicates => "1:rep+1" (e.g. for 3 replicates: "1:4")
# can be flexibly changed

ins = 6
rep = 3


# specify how many inserts are "blanks" and how many "cells"
  
n.blanks <- 3
n.cells <- 3

# set the working directory
# the excel template file MUST be stored here
# when copy file path from Windows Explorer please replace "\" with "//" and "\\" with "//"

wd <- "D://Hannah//"

# show the files available in the workind directory
list.files(wd)


# make list with excel sheets as elements
# read in an Excel workbook containing a sheet for each data point&direction (A->B or B->A)

wb <- loadWorkbook(file.path(wd,"20150409_LY_RHO_AB_BA.xlsx"))
ws = readWorksheet(wb, sheet = getSheets(wb))

summary(ws)

# STOP MODIFYING HERE :)))

###########################################################################################################################
### calibration curve

# extract the calibration measurement for RHO and LY
# ! in this setup RHO must be the first dye and LY the second
# the excel sheet with calibration data MUST be named "Calibration" (!no additional characters like invizible spaces:))) )
# calib. data for RHO starts at row 30 and for LY at row !!!62

rho.c <- as.data.frame(ws$Calibration[30:(30+ins+1),1:(rep+1)])
rho.c[,1] <- NULL
rho.c <- apply(rho.c,2,as.numeric)
colnames(rho.c) <- paste(rep("Repl",rep),rep(1:rep),sep="")
rownames(rho.c) <- paste(rep("Ins",(ins+2)),LETTERS[1:(ins+2)],sep="")

ly.c <- as.data.frame(ws$Calibration[62:(62+ins+1),1:(rep+1)])
ly.c[,1] <- NULL
ly.c <- apply(ly.c,2,as.numeric)
colnames(ly.c) <- paste(rep("Repl",rep),rep(1:rep),sep="")
rownames(ly.c) <- paste(rep("Ins",(ins+2)),LETTERS[1:(ins+2)],sep="")

# calculate average fluo. for each insert and dye
av.rho.c <- rowMeans(rho.c, na.rm=TRUE)
av.ly.c <- rowMeans(ly.c, na.rm=TRUE)


# specify the "dilution" series
# hard code the last entry to "0"

rho.c.f <- C0.ab.rho/2^(seq(0,(ins+1),1))
rho.c.f[(ins+2)] <- 0


ly.c.f <- C0.ab.ly/2^(seq(0,(ins+1),1))
ly.c.f[(ins+2)] <- 0


# calculate and plot the calibration curves for each dye
# save plot as pdf (in the working directory)

dev.off()
pdf(file.path(paste(wd,"CalibC.pdf",sep="")),width=11, height=8.5,pointsize=12, paper='special')
par(mfrow=c(1,2),oma=c(0,0,2,0))

plot(av.rho.c ~ rho.c.f, pch=18, col="gray",xlab="Conc. [ug/ml]",ylab="Average Fluo.",main="RHO")

lm_rho.c <- round(coef(lm(av.rho.c ~ rho.c.f)),2)
lm_rho.c

mtext(bquote(y == .(lm_rho.c[2])*x + .(lm_rho.c[1])), adj=1, padj=0) 
abline(lm(av.rho.c ~ rho.c.f),col="magenta",lwd=0.5)

plot(av.ly.c ~ ly.c.f, pch=18, col="gray",xlab="Conc. [ug/ml]",ylab="Average Fluo.",main="LY")

lm_ly.c <- round(coef(lm(av.ly.c ~ ly.c.f)),2)
lm_ly.c

mtext(bquote(y == .(lm_ly.c[2])*x + .(lm_ly.c[1])), adj=1, padj=0) 
abline(lm(av.ly.c ~ ly.c.f),col="magenta",lwd=0.5)
title("Calibration Curves", outer=TRUE)
dev.off()

# extract the fluo. data for each time point (t05 to t4) and direction (A to B and B to A) 
# store as list

###############################################################################################################
# RHO

pdf(file.path(paste(wd,"RHO_CalibratedFluo.pdf",sep="")),width=11, height=8.5,pointsize=12, paper='special')
par(mfrow=c(1,2),oma=c(0,0,2,0))
fluo.RHO <- list()

for (fl in grep("Calibration",names(ws),invert=TRUE,value = TRUE)){# all time points and directions except for "Calibration"
  
  # extract the fluo. data for RHO
  fl.datRHO <- as.data.frame(ws[[fl]][30:(30+ins-1),1:(rep+1)])
  fl.datRHO[,1] <- NULL
  fl.datRHO <- apply(fl.datRHO,2,as.numeric)
  colnames(fl.datRHO) <- paste(rep("Repl",rep),rep(1:rep),sep="")

  # in this setup "blanks" MUST be located in the excel file BEFORE "cells"
  rownames(fl.datRHO) <- rep("Ins",ins)
  rownames(fl.datRHO)[1:n.blanks] <- paste(rep("Ins",(n.blanks)),LETTERS[1:(n.blanks)],rep("_blank",n.blanks),sep="")
  rownames(fl.datRHO)[(n.blanks+1):(n.blanks+n.cells)] <- paste(rep("Ins",(n.cells)),LETTERS[(n.blanks+1):(n.blanks+n.cells)],rep("_cells",n.cells),sep="")
  fl.datRHO <- as.data.frame(fl.datRHO)
  
  # calibrate each replicate and then calculate the average and SD
  # add columns for calib. replicates
  
  for (i in 1:rep){
    fl.datRHO <- cbind(fl.datRHO, (fl.datRHO[,i] - lm_rho.c[1])/lm_rho.c[2])
  }
  colnames(fl.datRHO) <- c(paste(rep("Repl",rep),rep(1:rep),sep=""),paste("calib_Repl",rep(1:rep),sep=""))

  # calculate and add columns with replicates average and SD (of calibrated fluo.RHO. data) 
  fl.datRHO$av.fl.datRHO <- rowMeans(fl.datRHO[,grep("^calib*",colnames(fl.datRHO))], na.rm=TRUE)
  fl.datRHO$sd.fl.datRHO <- apply(fl.datRHO[,grep("^calib*",colnames(fl.datRHO))],1, function(x) sd(x, na.rm=TRUE))
  
  fluo.RHO[[fl]] <- fl.datRHO

  boxplot(cbind(t(fl.datRHO[grep("*blank",rownames(fl.datRHO)),grep("calib*",colnames(fl.datRHO))]),
              t(fl.datRHO[grep("*cells",rownames(fl.datRHO)),grep("calib*",colnames(fl.datRHO))])),
        col=c(rep("gray",n.blanks),rep("green",n.cells)),xaxt="n",
        xlab=paste("replicates ", fl,sep=""), ylab="Calibrated fluo.RHO.")

  # in this setup n.blanks MUST be equal to n.cells
  axis(1, cex.axis=0.75, las=1, at=seq(1,ins), labels=paste(rep("Ins.",ins),LETTERS[1:n.blanks],sep=""))
  legend("topright",c("blanks","cells"),fill=c("gray","green"),bty="n")

  errbar(c(1:ins),fl.datRHO$av.fl.datRHO,fl.datRHO$av.fl.datRHO + fl.datRHO$sd.fl.datRHO, 
       fl.datRHO$av.fl.datRHO - fl.datRHO$sd.fl.datRHO,
       xaxt="n", xlab=paste("Replicate Average and SD ",fl,sep=""),ylab="Calibrated fluo.RHO.",
       type="p", pch=18, col=c(rep("gray",n.blanks),rep("green",n.cells)))
  axis(1, cex.axis=0.75, las=1, at=seq(1,ins), labels=paste(rep("Ins.",ins),LETTERS[1:n.blanks],sep=""))
  legend("topright",c("blanks","cells"),col=c("gray","green"),pch=18,bty="n")

  title("RHO Calibrated fluorescence per time point and direction (A to B or B to A)", outer=TRUE)

}
dev.off()

for (i in 1:(length(names(fluo.RHO)))){
  names(fluo.RHO)[i] <- paste(names(fluo.RHO)[i],"Calib_AV_SD",sep="_")
}

summary(fluo.RHO)
names(fluo.RHO)

############################################################################################################
# LY

pdf(file.path(paste(wd,"LY_CalibratedFluo.pdf",sep="")),width=11, height=8.5,pointsize=12, paper='special')
par(mfrow=c(1,2),oma=c(0,0,2,0))
fluo.LY <- list()

for (fl in grep("Calibration",names(ws),invert=TRUE,value = TRUE)){# all time points and directions except for "Calibration"
  
  # extract the fluo. data for LY
  fl.datLY <- as.data.frame(ws[[fl]][60:(60+ins-1),1:(rep+1)])
  fl.datLY[,1] <- NULL
  fl.datLY <- apply(fl.datLY,2,as.numeric)
  colnames(fl.datLY) <- paste(rep("Repl",rep),rep(1:rep),sep="")
  
  # in this setup "blanks" MUST be located in the excel file BEFORE "cells"
  rownames(fl.datLY) <- rep("Ins",ins)
  rownames(fl.datLY)[1:n.blanks] <- paste(rep("Ins",(n.blanks)),LETTERS[1:(n.blanks)],rep("_blank",n.blanks),sep="")
  rownames(fl.datLY)[(n.blanks+1):(n.blanks+n.cells)] <- paste(rep("Ins",(n.cells)),LETTERS[(n.blanks+1):(n.blanks+n.cells)],rep("_cells",n.cells),sep="")
  fl.datLY <- as.data.frame(fl.datLY)
  
  # calibrate each replicate and then calculate the average and SD
  # add columns for calib. replicates
  
  for (i in 1:rep){
    fl.datLY <- cbind(fl.datLY, (fl.datLY[,i] - lm_ly.c[1])/lm_ly.c[2])
  }
  colnames(fl.datLY) <- c(paste(rep("Repl",rep),rep(1:rep),sep=""),paste("calib_Repl",rep(1:rep),sep=""))
  
  # calculate and add columns with replicates average and SD (of calibrated fluo.LY. data) 
  fl.datLY$av.fl.datLY <- rowMeans(fl.datLY[,grep("^calib*",colnames(fl.datLY))], na.rm=TRUE)
  fl.datLY$sd.fl.datLY <- apply(fl.datLY[,grep("^calib*",colnames(fl.datLY))],1, function(x) sd(x, na.rm=TRUE))
  
  fluo.LY[[fl]] <- fl.datLY
  
  boxplot(cbind(t(fl.datLY[grep("*blank",rownames(fl.datLY)),grep("calib*",colnames(fl.datLY))]),
                t(fl.datLY[grep("*cells",rownames(fl.datLY)),grep("calib*",colnames(fl.datLY))])),
          col=c(rep("gray",n.blanks),rep("green",n.cells)),xaxt="n",
          xlab=paste("replicates ", fl,sep=""), ylab="Calibrated fluo.LY.")
  
  # in this setup n.blanks MUST be equal to n.cells
  axis(1, cex.axis=0.75, las=1, at=seq(1,ins), labels=paste(rep("Ins.",ins),LETTERS[1:n.blanks],sep=""))
  legend("topright",c("blanks","cells"),fill=c("gray","green"),bty="n")
  
  errbar(c(1:ins),fl.datLY$av.fl.datLY,fl.datLY$av.fl.datLY + fl.datLY$sd.fl.datLY, 
         fl.datLY$av.fl.datLY - fl.datLY$sd.fl.datLY,
         xaxt="n", xlab=paste("Replicate Average and SD ",fl,sep=""),ylab="Calibrated fluo.LY.",
         type="p", pch=18, col=c(rep("gray",n.blanks),rep("green",n.cells)))
  axis(1, cex.axis=0.75, las=1, at=seq(1,ins), labels=paste(rep("Ins.",ins),LETTERS[1:n.blanks],sep=""))
  legend("topright",c("blanks","cells"),col=c("gray","green"),pch=18,bty="n")
  
  title("LY Calibrated fluorescence per time point and direction (A to B or B to A)", outer=TRUE)
  
}
dev.off()

for (i in 1:(length(names(fluo.LY)))){
  names(fluo.LY)[i] <- paste(names(fluo.LY)[i],"Calib_AV_SD",sep="_")
}

summary(fluo.LY)
names(fluo.LY)

################################################################################################################

# Mass over time for each dye and direction

################
# RHO & A to B #
################

summary(fluo.RHO)
fluo.RHO[[1]]

blanks.l <- list()
for (i in grep("^A",names(fluo.RHO),value = TRUE)){  
  bl <- as.data.frame(fluo.RHO[[i]][grep("*blank",rownames(fluo.RHO[[i]])),"av.fl.datRHO"])
  blanks.l[[i]] <- bl
}


blanks <- do.call("cbind",blanks.l)
rownames(blanks) <- paste(rep("Ins",(n.blanks)),LETTERS[1:(n.blanks)],rep("_blank",n.blanks),sep="")
colnames(blanks) <- paste(grep("^A",names(fluo.RHO),value=TRUE))
colnames(blanks) <- gsub("A-B ","",colnames(blanks))
colnames(blanks) <- gsub("_Calib_AV_SD","",colnames(blanks))
blanks <- blanks[,order(colnames(blanks),decreasing=FALSE)]

# adjust Volume per time point
adjVol <- seq(V.ab.brec, by=-V,length.out=ncol(blanks))

# Calibrated Quantities to Mass 
blanks.mass <- t(apply(as.matrix(blanks),1, function(x) adjVol * x))
# replace negative values with 0
blanks.mass[which(blanks.mass<0)] <- 0
blanks.mass

#############################################################################################

cells.l <- list()
for (i in grep("^A",names(fluo.RHO),value = TRUE)){  
  bl <- as.data.frame(fluo.RHO[[i]][grep("*cells",rownames(fluo.RHO[[i]])),"av.fl.datRHO"])
  cells.l[[i]] <- bl
}


cells <- do.call("cbind",cells.l)
rownames(cells) <- paste(rep("Ins",(n.cells)),LETTERS[1:(n.cells)],rep("_cells",n.cells),sep="")
colnames(cells) <- paste(grep("^A",names(fluo.RHO),value=TRUE))
colnames(cells) <- gsub("A-B ","",colnames(cells))
colnames(cells) <- gsub("_Calib_AV_SD","",colnames(cells))
cells <- cells[,order(colnames(cells),decreasing=FALSE)]

# adjust Volume per time point
adjVol <- seq(V.ab.brec, by=-V,length.out=ncol(cells))
adjVol
# Calibrated Quantities to Mass 
cells.mass <- t(apply(as.matrix(cells),1, function(x) adjVol * x))
# replace negative values with 0
cells.mass[which(cells.mass<0)] <- 0
cells.mass

#################################################################################################
pdf(file.path(paste(wd,"RHO_MassOverTime_AtoB.pdf",sep="")),width=11, height=8.5,pointsize=12, paper='special')
par(mfrow=c(1,3),oma=c(0,0,2,0))


mass.l <- list()
for (i in 1:(ins/2)){
  mass.l[[i]] <- rbind(blanks.mass[i,],cells.mass[i,])
  rownames(mass.l[[i]]) <- c(rownames(blanks.mass)[i],rownames(cells.mass)[i])
}

names(mass.l) <- paste(rep("Ins",(ins/2)),LETTERS[1:(ins/2)],sep="")


for (i in names(mass.l)){
  xLabLocs <- plot(mass.l[[i]][1,] ~ seq(1,ncol(mass.l[[i]])), type="b", pch=18, 
                   col="gray",xaxt='n',xlab="time (h)",ylab="Mass",main=paste(gsub("_blank","",rownames(mass.l[[i]])[1])))
  axis(1, cex.axis=0.75, las=1, at=seq(1,ncol(blanks.mass)), labels=colnames(blanks.mass))
  lines(mass.l[[i]][2,] ~ seq(1,ncol(mass.l[[i]])) , pch=18, col="green",type="b")
  legend("topleft",c("blanks","cells"),fill=c("gray","green"))
}


title("Inserts ¦ A to B", outer=TRUE)

dev.off()



# Calculate Permeability

summary(mass.l)
mass.l[[1]]

perm.l <- list()
for(i in names(mass.l)){
  lm_blanks <- coef(lm(mass.l[[i]][1,] ~ seq(1,ncol(mass.l[[i]]))))
  lm_cells <- coef(lm(mass.l[[i]][2,] ~ seq(1,ncol(mass.l[[i]]))))
  perm.blanks <- as.numeric(lm_blanks[2]) * as.numeric(1/A*C0.ab.rho)
  perm.cells <- as.numeric(lm_cells[2]) * as.numeric(1/A*C0.ab.rho)
  perm.l[[i]] <- cbind(perm.blanks,perm.cells)
}

summary(perm.l)
perm.l[[1]]

perm.rho.ab <- do.call("rbind",perm.l)
rownames(perm.rho.ab) <- names(perm.l)
perm.rho.ab

################
# RHO & B to A #
################

summary(fluo.RHO)
fluo.RHO[[1]]

blanks.l <- list()
for (i in grep("^B",names(fluo.RHO),value = TRUE)){  
  bl <- as.data.frame(fluo.RHO[[i]][grep("*blank",rownames(fluo.RHO[[i]])),"av.fl.datRHO"])
  blanks.l[[i]] <- bl
}

summary(blanks.l)

blanks <- do.call("cbind",blanks.l)
rownames(blanks) <- paste(rep("Ins",(n.blanks)),LETTERS[1:(n.blanks)],rep("_blank",n.blanks),sep="")
colnames(blanks) <- paste(grep("^B",names(fluo.RHO),value=TRUE))
colnames(blanks) <- gsub("B-A ","",colnames(blanks))
colnames(blanks) <- gsub("_Calib_AV_SD","",colnames(blanks))
blanks <- blanks[,order(colnames(blanks),decreasing=FALSE)]
blanks
# adjust Volume per time point
adjVol <- seq(V.ba.arec, by=-V,length.out=ncol(blanks))
adjVol
# Calibrated Quantities to Mass 
blanks.mass <- t(apply(as.matrix(blanks),1, function(x) adjVol * x))
# replace negative values with 0
blanks.mass[which(blanks.mass<0)] <- 0
blanks.mass

#############################################################################################

cells.l <- list()
for (i in grep("^B",names(fluo.RHO),value = TRUE)){  
  bl <- as.data.frame(fluo.RHO[[i]][grep("*cells",rownames(fluo.RHO[[i]])),"av.fl.datRHO"])
  cells.l[[i]] <- bl
}


cells <- do.call("cbind",cells.l)
rownames(cells) <- paste(rep("Ins",(n.cells)),LETTERS[1:(n.cells)],rep("_cells",n.cells),sep="")
colnames(cells) <- paste(grep("^B",names(fluo.RHO),value=TRUE))
colnames(cells) <- gsub("B-A ","",colnames(cells))
colnames(cells) <- gsub("_Calib_AV_SD","",colnames(cells))
cells <- cells[,order(colnames(cells),decreasing=FALSE)]

# adjust Volume per time point
adjVol <- seq(V.ba.arec, by=-V,length.out=ncol(cells))
adjVol
# Calibrated Quantities to Mass 
cells.mass <- t(apply(as.matrix(cells),1, function(x) adjVol * x))
# replace negative values with 0
cells.mass[which(cells.mass<0)] <- 0
cells.mass

######################################################################################################################################
pdf(file.path(paste(wd,"RHO_MassOverTime_BtoA.pdf",sep="")),width=11, height=8.5,pointsize=12, paper='special')
par(mfrow=c(1,3),oma=c(0,0,2,0))


mass.l <- list()
for (i in 1:(ins/2)){
  mass.l[[i]] <- rbind(blanks.mass[i,],cells.mass[i,])
  rownames(mass.l[[i]]) <- c(rownames(blanks.mass)[i],rownames(cells.mass)[i])
}

names(mass.l) <- paste(rep("Ins",(ins/2)),LETTERS[1:(ins/2)],sep="")


for (i in names(mass.l)){
  xLabLocs <- plot(mass.l[[i]][1,] ~ seq(1,ncol(mass.l[[i]])), type="b", pch=18, 
                   col="gray",xaxt='n',xlab="time (h)",ylab="Mass",main=paste(gsub("_blank","",rownames(mass.l[[i]])[1])))
  axis(1, cex.axis=0.75, las=1, at=seq(1,ncol(blanks.mass)), labels=colnames(blanks.mass))
  lines(mass.l[[i]][2,] ~ seq(1,ncol(mass.l[[i]])) , pch=18, col="green",type="b")
  legend("topleft",c("blanks","cells"),fill=c("gray","green"))
}


title("Inserts ¦ B to A", outer=TRUE)

dev.off()

# Calculate Permeability

summary(mass.l)
mass.l[[1]]

perm.l <- list()
for(i in names(mass.l)){
  lm_blanks <- coef(lm(mass.l[[i]][1,] ~ seq(1,ncol(mass.l[[i]]))))
  lm_cells <- coef(lm(mass.l[[i]][2,] ~ seq(1,ncol(mass.l[[i]]))))
  perm.blanks <- as.numeric(lm_blanks[2]) * as.numeric(1/A*C0.ba.rho)
  perm.cells <- as.numeric(lm_cells[2]) * as.numeric(1/A*C0.ba.rho)
  perm.l[[i]] <- cbind(perm.blanks,perm.cells)
}

summary(perm.l)
perm.l[[1]]

perm.rho.ba <- do.call("rbind",perm.l)
rownames(perm.rho.ba) <- names(perm.l)
perm.rho.ba

#######################################################################################################################################
################
# LY & A to B #
################

summary(fluo.LY)
fluo.LY[[1]]

blanks.l <- list()
for (i in grep("^A",names(fluo.LY),value = TRUE)){  
  bl <- as.data.frame(fluo.LY[[i]][grep("*blank",rownames(fluo.LY[[i]])),"av.fl.datLY"])
  blanks.l[[i]] <- bl
}


blanks <- do.call("cbind",blanks.l)
rownames(blanks) <- paste(rep("Ins",(n.blanks)),LETTERS[1:(n.blanks)],rep("_blank",n.blanks),sep="")
colnames(blanks) <- paste(grep("^A",names(fluo.LY),value=TRUE))
colnames(blanks) <- gsub("A-B ","",colnames(blanks))
colnames(blanks) <- gsub("_Calib_AV_SD","",colnames(blanks))
blanks <- blanks[,order(colnames(blanks),decreasing=FALSE)]

# adjust Volume per time point
adjVol <- seq(V.ab.brec, by=-V,length.out=ncol(blanks))
adjVol
# Calibrated Quantities to Mass 
blanks.mass <- t(apply(as.matrix(blanks),1, function(x) adjVol * x))
# replace negative values with 0
blanks.mass[which(blanks.mass<0)] <- 0
blanks.mass

#############################################################################################

cells.l <- list()
for (i in grep("^A",names(fluo.LY),value = TRUE)){  
  bl <- as.data.frame(fluo.LY[[i]][grep("*cells",rownames(fluo.LY[[i]])),"av.fl.datLY"])
  cells.l[[i]] <- bl
}


cells <- do.call("cbind",cells.l)
rownames(cells) <- paste(rep("Ins",(n.cells)),LETTERS[1:(n.cells)],rep("_cells",n.cells),sep="")
colnames(cells) <- paste(grep("^A",names(fluo.LY),value=TRUE))
colnames(cells) <- gsub("A-B ","",colnames(cells))
colnames(cells) <- gsub("_Calib_AV_SD","",colnames(cells))
cells <- cells[,order(colnames(cells),decreasing=FALSE)]

# adjust Volume per time point
adjVol <- seq(V.ab.brec, by=-V,length.out=ncol(cells))
adjVol
# Calibrated Quantities to Mass 
cells.mass <- t(apply(as.matrix(cells),1, function(x) adjVol * x))
# replace negative values with 0
cells.mass[which(cells.mass<0)] <- 0
cells.mass

#################################################################################################
pdf(file.path(paste(wd,"LY_MassOverTime_AtoB.pdf",sep="")),width=11, height=8.5,pointsize=12, paper='special')
par(mfrow=c(1,3),oma=c(0,0,2,0))


mass.l <- list()
for (i in 1:(ins/2)){
  mass.l[[i]] <- rbind(blanks.mass[i,],cells.mass[i,])
  rownames(mass.l[[i]]) <- c(rownames(blanks.mass)[i],rownames(cells.mass)[i])
}

names(mass.l) <- paste(rep("Ins",(ins/2)),LETTERS[1:(ins/2)],sep="")


for (i in names(mass.l)){
  xLabLocs <- plot(mass.l[[i]][1,] ~ seq(1,ncol(mass.l[[i]])), type="b", pch=18, 
                   col="gray",xaxt='n',xlab="time (h)",ylab="Mass",main=paste(gsub("_blank","",rownames(mass.l[[i]])[1])))
  axis(1, cex.axis=0.75, las=1, at=seq(1,ncol(blanks.mass)), labels=colnames(blanks.mass))
  lines(mass.l[[i]][2,] ~ seq(1,ncol(mass.l[[i]])) , pch=18, col="green",type="b")
  legend("topleft",c("blanks","cells"),fill=c("gray","green"))
}


title("Inserts ¦ A to B", outer=TRUE)

dev.off()



# Calculate Permeability

summary(mass.l)
mass.l[[1]]

perm.l <- list()
for(i in names(mass.l)){
  lm_blanks <- coef(lm(mass.l[[i]][1,] ~ seq(1,ncol(mass.l[[i]]))))
  lm_cells <- coef(lm(mass.l[[i]][2,] ~ seq(1,ncol(mass.l[[i]]))))
  perm.blanks <- as.numeric(lm_blanks[2]) * as.numeric(1/A*C0.ab.ly)
  perm.cells <- as.numeric(lm_cells[2]) * as.numeric(1/A*C0.ab.ly)
  perm.l[[i]] <- cbind(perm.blanks,perm.cells)
}

summary(perm.l)
perm.l[[1]]

perm.ly.ab <- do.call("rbind",perm.l)
rownames(perm.ly.ab) <- names(perm.l)
perm.ly.ab

################
# LY & B to A #
################

summary(fluo.LY)
fluo.LY[[1]]

blanks.l <- list()
for (i in grep("^B",names(fluo.LY),value = TRUE)){  
  bl <- as.data.frame(fluo.LY[[i]][grep("*blank",rownames(fluo.LY[[i]])),"av.fl.datLY"])
  blanks.l[[i]] <- bl
}

summary(blanks.l)

blanks <- do.call("cbind",blanks.l)
rownames(blanks) <- paste(rep("Ins",(n.blanks)),LETTERS[1:(n.blanks)],rep("_blank",n.blanks),sep="")
colnames(blanks) <- paste(grep("^B",names(fluo.LY),value=TRUE))
colnames(blanks) <- gsub("B-A ","",colnames(blanks))
colnames(blanks) <- gsub("_Calib_AV_SD","",colnames(blanks))
blanks <- blanks[,order(colnames(blanks),decreasing=FALSE)]
blanks
# adjust Volume per time point
adjVol <- seq(V.ba.arec, by=-V,length.out=ncol(blanks))
adjVol
# Calibrated Quantities to Mass 
blanks.mass <- t(apply(as.matrix(blanks),1, function(x) adjVol * x))
# replace negative values with 0
blanks.mass[which(blanks.mass<0)] <- 0
blanks.mass

#############################################################################################

cells.l <- list()
for (i in grep("^B",names(fluo.LY),value = TRUE)){  
  bl <- as.data.frame(fluo.LY[[i]][grep("*cells",rownames(fluo.LY[[i]])),"av.fl.datLY"])
  cells.l[[i]] <- bl
}


cells <- do.call("cbind",cells.l)
rownames(cells) <- paste(rep("Ins",(n.cells)),LETTERS[1:(n.cells)],rep("_cells",n.cells),sep="")
colnames(cells) <- paste(grep("^B",names(fluo.LY),value=TRUE))
colnames(cells) <- gsub("B-A ","",colnames(cells))
colnames(cells) <- gsub("_Calib_AV_SD","",colnames(cells))
cells <- cells[,order(colnames(cells),decreasing=FALSE)]

# adjust Volume per time point
adjVol <- seq(V.ba.arec, by=-V,length.out=ncol(cells))
adjVol
# Calibrated Quantities to Mass 
cells.mass <- t(apply(as.matrix(cells),1, function(x) adjVol * x))
# replace negative values with 0
cells.mass[which(cells.mass<0)] <- 0
cells.mass

######################################################################################################################################
pdf(file.path(paste(wd,"LY_MassOverTime_BtoA.pdf",sep="")),width=11, height=8.5,pointsize=12, paper='special')
par(mfrow=c(1,3),oma=c(0,0,2,0))


mass.l <- list()
for (i in 1:(ins/2)){
  mass.l[[i]] <- rbind(blanks.mass[i,],cells.mass[i,])
  rownames(mass.l[[i]]) <- c(rownames(blanks.mass)[i],rownames(cells.mass)[i])
}

names(mass.l) <- paste(rep("Ins",(ins/2)),LETTERS[1:(ins/2)],sep="")


for (i in names(mass.l)){
  xLabLocs <- plot(mass.l[[i]][1,] ~ seq(1,ncol(mass.l[[i]])), type="b", pch=18, 
                   col="gray",xaxt='n',xlab="time (h)",ylab="Mass",main=paste(gsub("_blank","",rownames(mass.l[[i]])[1])))
  axis(1, cex.axis=0.75, las=1, at=seq(1,ncol(blanks.mass)), labels=colnames(blanks.mass))
  lines(mass.l[[i]][2,] ~ seq(1,ncol(mass.l[[i]])) , pch=18, col="green",type="b")
  legend("topleft",c("blanks","cells"),fill=c("gray","green"))
}


title("Inserts ¦ B to A", outer=TRUE)

dev.off()


# Calculate Permeability

summary(mass.l)
mass.l[[1]]

perm.l <- list()
for(i in names(mass.l)){
  lm_blanks <- coef(lm(mass.l[[i]][1,] ~ seq(1,ncol(mass.l[[i]]))))
  lm_cells <- coef(lm(mass.l[[i]][2,] ~ seq(1,ncol(mass.l[[i]]))))
  perm.blanks <- as.numeric(lm_blanks[2]) * as.numeric(1/A*C0.ba.ly)
  perm.cells <- as.numeric(lm_cells[2]) * as.numeric(1/A*C0.ba.ly)
  perm.l[[i]] <- cbind(perm.blanks,perm.cells)
}

summary(perm.l)
perm.l[[1]]

perm.ly.ba <- do.call("rbind",perm.l)
rownames(perm.ly.ba) <- names(perm.l)
perm.ly.ba


#######################################################################################################################################

#save excel with multiple sheets with perm. for all inserts and directions

perm.all <- list(perm.rho.ab,perm.rho.ba,perm.ly.ab,perm.ly.ba)
perm.all[[4]]
summary(perm.all)
names(perm.all) <- c("perm.rho.ab","perm.rho.ba","perm.ly.ab","perm.ly.ba")

wb.res <- loadWorkbook(file.path(paste(wd,"Perm.xlsx",sep="")), create = TRUE)
createSheet(wb.res, name=names(perm.all))
writeWorksheet(wb.res, perm.all,names(perm.all),header=TRUE,rownames="Ins")
saveWorkbook(wb.res)


