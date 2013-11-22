These 3 short VB macros were written to solve a specific problem my friend was having with randomizing seriesid slides for a research project.

The goal was to create 3 different macros for 3 different Power Points.

**First - Total Randomize**
Create a macro that when executed will randomize all of the slides in a ppt file.

**Second - 3 Series Randomize**
In a PowerPoint of 30 slides there are 10 groups of 3 slides wherein the first slide is always static, the following four are to be randomized and each group of 3 (kept intact) are then randomized.
I.E.
Pre-randomize **[1]**[2][3]-**[4]**[5][6]-**[7]**[8][9]  
Post-randomized **[4]**[5][6]-**[1]**[3][2]-**[7]**[9][8]

This concept will be explained further in the next section.

**Third - 5 Series Randomize**
Similar to the 3 Series Randomize - In a PowerPoint of 50 slides there are 10 groups of 5 slides wherein the first slide is always static, the following four are to be randomized and each group of 5 (kept intact) are then randomized.

In completion this script should work in two loops. The first loop will keep alpha slide static and then proceed to randomize the following 4 slides.

![Alt text](http://hashtagnerd.com/wp-content/uploads/2012/11/VB5set_drawing1.jpg "First Loop")


![Alt text](http://hashtagnerd.com/wp-content/uploads/2012/11/VB5set_drawing2.jpg "Second Loop")



**Problems I Faced**

1 – There are no distinguishing names for slides in PowerPoint other then the position that they current exist. This causes a problem when trying to randomize a slide to somewhere else within a PowerPoint while keeping its other 4 slides in coherency.

2 – VB cannot support selecting multiple slides and moving them together. In the way you might think of selecting several slides by clicking, holding down shift and selecting another slide to grab a series of slides.

3 – When using VB you can either use the Cut and Paste method or you can use the MoveTo method in which you set the slide to move and the location it is to move. The problem or rather I should say the stumbling block is understanding that when you do either of these methods you must account for the fact that when a slide is moved the slide that it is moving to is going to shift down. Meaning that the location that you plan to move the slide must be offset in this consideration.

4 - A short time table to learn VB as it applies in Power Point left me to create a functional but tailor made piece of code. 
