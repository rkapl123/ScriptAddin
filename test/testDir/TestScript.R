library(ggplot2)
initial.options <- commandArgs(trailingOnly = FALSE)
script.name <- sub("--file=", "", initial.options[grep("--file=", initial.options)])
script.basename <- dirname(script.name)
if (length(script.basename) == 0) {
  # this only works RGui
  script.basename <- getSrcDirectory(function(x) {x})
}
setwd(script.basename)

#test_in <-as.matrix(read.csv2("test_in.txt",sep="\t",header=FALSE))
#mode(test_in) <-"numeric"
test_in <- read.csv2("test_in.txt",dec=".",sep="\t",header=TRUE)
write.table(test_in[,1],file="test_out.txt",sep="\t",row.names=FALSE,col.names=FALSE,quote=FALSE)

gplot <- ggplot(data = test_in, aes(x = in2, y = in1) ) + geom_point()
png(filename="testdiagram.png")
print(gplot)
dev.off()

