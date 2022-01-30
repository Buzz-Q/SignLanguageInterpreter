The project is a proof of concept for a sign language interpretation using
3 flex sensors applied on the elbow, middle finger and thumb.

dataAquisition2 File: A python program to connect python to arduino and 
read the 3 flex sensors data when the signer begins to sign.
The thresholds are adjustable according to the sensors' default.

fixingZeros File: A python program to clean the data from outliers.
The thresholds can be adjusted according to the data which can be viwed on wordPlot

wordPlot File: a Matlab file to plot a sample of the data collected.

finalMODELS.ipynb File: A python program on anaconda IDE
containing the final machine learning models for word prediction.

