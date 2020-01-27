a = ['1','2','3','4','5','6']
print len(a)
for i in range(0,len(a)):
    value = i%3
    print value
    if value==0:
        print "This is Company"
    elif value==1:
        print "this is Designation"
    else:
        print "this is year"