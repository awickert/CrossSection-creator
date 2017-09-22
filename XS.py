import openpyxl
import numpy as np
from matplotlib import pyplot as plt
from scipy.interpolate import interp1d

def import_sheet(file_name, sheet_name):
  wb = openpyxl.load_workbook(file_name)
  sheet = wb.get_sheet_by_name(sheet_name)
  return sheet

def values_in_column(sheet, n, m=0):
  outlist = []
  cells = sheet.columns[n][m:]
  for cell in cells:
    value = cell.value
    if value:
      outlist.append(value)
  outarray = np.array(outlist)
  return outarray

allData = {}
file_name = 'Gavilan_CC.xlsx'
wb = openpyxl.load_workbook(file_name)
for sheet in wb:
  xcol = 0
  ycol = 1
  zcol = 2
  x = values_in_column(sheet, xcol, 1)
  y = values_in_column(sheet, xcol, 1)
  z = values_in_column(sheet, zcol, 1)
  dist = ( (x-x[0])**2 + (y - y[0])**2 )**0.5
  #allData.append(np.array([dist,x,y,z]))
  allData[sheet.title] = {'dist':dist, 'e':x, 'n':y, 'z':z}

# Distance range
dist_all = []
for layerName in allData:
  dist_all += list(allData[layerName]['dist'])
dist_min = np.floor(np.min(dist_all))
dist_max = np.ceil(np.max(dist_all))

# Elevation range
z_all = []
for layerName in allData:
  z_all += list(allData[layerName]['z'])
z_min = np.floor(np.min(z_all))
z_max = np.ceil(np.max(z_all))

# Interpolate
dist_interp = np.arange(dist_min, dist_max+1)
z_interp = {}
for layerName in allData:
  f = interp1d(allData[layerName]['dist'], allData[layerName]['z'], bounds_error=False, fill_value=np.nan) #fill_value='extrapolate')
  z_interp[layerName] = f(dist_interp)
  print f(dist_interp)
  
# Truncate at surface
for layerName in z_interp:
  z_interp[layerName][z_interp[layerName] > z_interp['GroundSurface']] = z_interp['GroundSurface'][z_interp[layerName] > z_interp['GroundSurface']]

# Create layers -- Rachel gave bottom
layer_bottoms = z_interp.copy()
del layer_bottoms['GroundSurface']
layer_tops = {}
for layerName in layer_bottoms:
  OtherSurfaces = []
  for layerName2 in z_interp:
    if layerName2 != layerName:
      OtherSurfaces.append(z_interp[layerName2])
  OtherSurfaces = np.vstack(OtherSurfaces)
  # Add a large number to those that are below to make them not the minimum
  OtherSurfaces[OtherSurfaces < layer_bottoms[layerName]] += 1E3
  layer_tops[layerName] = np.nanmin(OtherSurfaces, axis=0)

fill_colors = {}
fill_colors['CH1'] = 'orange'
fill_colors['CH2'] = 'yellow'
fill_colors['LGM'] = 'mediumblue'
fill_colors['YD'] = 'lightskyblue'
fill_colors['YDLake'] = 'steelblue'
fill_colors['NG'] = 'indigo'
fill_colors['PFD'] = 'deeppink'
fill_colors['UVol'] = 'dimgray'
fill_colors['Basement'] = 'green'

plt.ion()
plt.figure(figsize=(12,4))
for layerName in layer_bottoms:
  plt.fill_between(dist_interp, layer_bottoms[layerName], layer_tops[layerName], facecolor=fill_colors[layerName], linewidth=1, label=layerName)
plt.plot(dist_interp, z_interp['GroundSurface'], 'k-', linewidth=2)
plt.xlim(dist_min, dist_max)
plt.ylim(np.ceil(z_min/10.)*10, np.ceil(z_max/100.)*100)
plt.legend(loc='lower left')
plt.xlabel('Distance along profile [m]', fontsize=16)
plt.ylabel('Elevation [m]', fontsize=16)
plt.tight_layout()

"""
# Just lines
plt.ion()
for layerName in z_interp:
  z = z_interp[layerName]
  plt.plot(dist_interp, z, '-', linewidth=2, label=layerName)
  plt.legend(loc='lower left')
"""

"""
# Original layer plotting
plt.ion()
for layerName in allData:
  values = allData[layerName]
  plt.plot(values['dist'], values['z'], '-', linewidth=2, label=layerName)
  plt.legend(loc='lower left')
"""
