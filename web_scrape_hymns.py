import pandas as pd
import json


# with open('songs.json') as f:
#   data = json.load(f)
#
#
# song_list = data['playlist']['list']
#
# song_number = []
# song_name = []
# for song in song_list:
#     print(song['songNumber'],song['name'],sep=',')
#     song_number.append(song['songNumber'])
#     song_name.append(song['name'])
#
# df = pd.DataFrame({'songNumber':song_number,'songName':song_name})
# df.to_csv('song.csv',index=False)
#



with open('sp_songs.json') as f:
  data = json.load(f)


song_list = data['playlist']['list']

song_number = []
song_name = []
for song in song_list:
    print(song['songNumber'],song['name'],sep=',')
    song_number.append(song['songNumber'])
    song_name.append(song['name'])

df = pd.DataFrame({'songNumber':song_number,'songName':song_name})
df.to_csv('song_sp.csv',index=False)
