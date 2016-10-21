import pygmap

unique_location = pygmap.cleanWB('C:/Users/windalfin_culmen/Documents/location.xlsx')
origin = 'alam sutera'
for i in range(len(unique_location)):
    pygmap.getDistance(origin, unique_location[i])
