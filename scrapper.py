# Import Module
import os
import re
import openrouteservice
from unidecode import unidecode
import xlsxwriter
import folium


openrouteservice_APIkey = '5b3ce3597851110001cf6248ab48e7c1a9584f30a793f97dda9afe07'
country = 'FR'
workplace_coordinates = [5.441204,43.230950] #CPPM
central_station_coordinates = [5.379806543690962, 43.30368401589245]  #st charles
bike_avg_speed = 15*1000/3600




# helper functions
def find_in_between(string,beginning,end,extremas_included):
    start_ind = string.find(beginning)
    end_ind = string[start_ind:].find(end) + start_ind
    if extremas_included:
        return string[start_ind:end_ind + len(end)]
    else:
        return string[start_ind + len(beginning) :end_ind]


# Folder Path
path = r"C:\Users\giova\Google Drive\FRANCE\Accomodation\WebsiteScrapingInfo\announces_pages"

# Change the directory
os.chdir(path)
# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook(path + '\comparisonTable.xlsx')

# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()


worksheet.write(0, 0, "announce number")
worksheet.write(0, 1, "website")
worksheet.write(0, 2, "link")
worksheet.write(0, 3, "Contact status")
worksheet.write(0, 4, "city & postcode")
worksheet.write(0, 5, "arrondisment")
worksheet.write(0, 6, "description")
worksheet.write(0, 7, "street")
worksheet.write(0, 8, "distance to work")
worksheet.write(0, 9, "time to work")
worksheet.write(0, 10, "distance to center")
worksheet.write(0, 11, "time to center")

worksheet.write(0, 12, "etage")
worksheet.write(0, 13, "balcon")
worksheet.write(0, 14, "cave")

worksheet.write(0, 15, "surface")
worksheet.write(0, 16, "pieces")
worksheet.write(0, 17, "price")
worksheet.write(0, 18, "price/m2")




apartment_global_index = 1


# iterate through all file
for file in os.listdir():
    # Check whether file is in text format or not
    if file.endswith(".html"):
        file_path = f"{path}\{file}"

        # call read text file function
        with open(file_path, 'r',encoding='utf8') as f:
            rawInfo = (f.read())

        # extract information
        #0 identify the website
        website = find_in_between(rawInfo, 'www.', '.com', False)
        worksheet.write(apartment_global_index, 1, website)

        if website == 'seloger':
            #1 get a number, global index
            announce_number = apartment_global_index
            worksheet.write(apartment_global_index, 0, announce_number)
            #2 get a link to the announce
            announce_link = find_in_between(rawInfo, 'https://', '.htm', True)
            worksheet.write(apartment_global_index, 2, announce_link)
            #3 a conctact colum that we will fill after
            announce_contact_status = ''
            worksheet.write(apartment_global_index, 3, announce_contact_status)
            #4 find the city and the postcode
            city_and_postcode = find_in_between(rawInfo,r'<span class="Localizationstyled__City',r'</span>', False).split('>')[1]
            announce_city = city_and_postcode.split()[0]
            announce_postcode = find_in_between(city_and_postcode,'(',')',False)
            worksheet.write(apartment_global_index, 4, announce_city + ' ' + announce_postcode)
            #5 arrondisment
            announce_arrondisment = str(divmod(int(announce_postcode),13000)[1])
            worksheet.write(apartment_global_index, 5, announce_arrondisment)
            #6 description
            announce_description = find_in_between(rawInfo,r'<div class="Descriptionstyled__StyledShowMoreText','</p>',False).split('<p>')[1]
            worksheet.write(apartment_global_index, 6, announce_description)
            #7 Find street
            # from html code
            hood_and_street = find_in_between(rawInfo,r'<h1 data-testid="gsl.uilib.Breadcrumb.LastElement">',r'</h1>',False).split(' - ')[1:3]
            announce_hood = hood_and_street[0]
            announce_street = hood_and_street[1]
            #attempt to match the description
            street_regexp = "[0-9]{1,3} .+ " + announce_street.split()[-1].lower()
            announce_street_v2 = re.findall(street_regexp, announce_description.lower())[0]
            announce_street_number = re.findall("[0-9]{1,3}",announce_street_v2)[0]
            worksheet.write(apartment_global_index, 7, announce_street)

            # 8 Find distance to workplace
            # 8a convert address into global coordinates
            client = openrouteservice.Client(key=openrouteservice_APIkey)
            res = openrouteservice.geocode.pelias_search(client,announce_street_v2,sources=['osm', 'oa', 'wof', 'gn'],country=country,circle_point=(5.383420,43.290300))
            announce_coordinates = res.get('features')[0].get('geometry').get('coordinates')

            # 8b find distance to work
            route_to_work = client.directions((tuple(announce_coordinates),tuple(workplace_coordinates)),profile='cycling-regular',radiuses=(100,100),preference='recommended')
            distance_to_work = route_to_work.get('routes')[0].get('summary').get('distance')
            time_to_work = route_to_work.get('routes')[0].get('summary').get('duration')
            honest_time_to_work = distance_to_work/bike_avg_speed
            # geometry = route_to_work['routes'][0]['geometry']
            # decoded = openrouteservice.convert.decode_polyline(geometry)
            # # Initialize the Map instance
            # m = folium.Map(location=announce_coordinates, zoom_start=10, control_scale=True,tiles="cartodbpositron")
            # folium.GeoJson(decoded).add_to(m)
            # m.save('route_map.html')

            worksheet.write(apartment_global_index, 8, distance_to_work)
            worksheet.write(apartment_global_index,9, honest_time_to_work)

            #8c find distance to saint charles
            route_to_center = client.directions((tuple(announce_coordinates), tuple(central_station_coordinates)),profile='cycling-regular', radiuses=(100, 100), preference='recommended')
            distance_to_center = route_to_center.get('routes')[0].get('summary').get('distance')
            time_to_center = route_to_center.get('routes')[0].get('summary').get('duration')
            honest_time_to_center = distance_to_center/bike_avg_speed

            worksheet.write(apartment_global_index, 10, distance_to_center)
            worksheet.write(apartment_global_index,11, honest_time_to_center)

            # 9  find the etage
            # find the word etage
            floor_keyword = "etage"
            search_res = re.findall(floor_keyword, unidecode(announce_description.lower()))
            if len(search_res) > 0:
                announce_floor = "1 or above"
            else:
                # find the word chaussée
                ground_floor_keyword = "chaussee"
                search_res = re.findall(ground_floor_keyword, unidecode(announce_description.lower()))
                if len(search_res) > 0:
                    announce_floor = "0"
                else:
                    announce_floor = "unknown"

            worksheet.write(apartment_global_index, 12, announce_floor)

            # 10 Balcon
            balcon_keyword = "balcon"
            search_res = re.findall(balcon_keyword, unidecode(announce_description.lower()))
            if len(search_res) > 0:
                announce_balcon = "yes"
            else:
                announce_balcon = "unknown"
            worksheet.write(apartment_global_index, 13, announce_balcon)
            # 11 Cave
            cave_keyword = "cave"
            search_res = re.findall(cave_keyword, unidecode(announce_description.lower()))
            if len(search_res) > 0:
                announce_cave = "yes"
            else:
                announce_cave = "unknown"
            worksheet.write(apartment_global_index, 14, announce_cave)
            # 12 surface
            surface_keyword = "[0-9]{1,3} m²"
            search_res = re.findall(surface_keyword, rawInfo)
            announce_surface = search_res[0].split()[0]
            worksheet.write(apartment_global_index, 15, announce_surface)
            # 13 pieces
            piece_keyword = "[0-9]{1,3} pièce"
            search_res = re.findall(piece_keyword, rawInfo)
            announce_pieces = search_res[0].split()[0]
            worksheet.write(apartment_global_index, 16, announce_pieces)
            # 14 price
            price_keyword = "[0-9]{1,3} € / mois"
            search_res = re.findall(price_keyword, rawInfo)
            announce_price_cc = search_res[0].split()[0]
            worksheet.write(apartment_global_index, 17, announce_price_cc)
            # 15 price per surface
            announce_price_per_m2 = float(announce_price_cc)/float(announce_surface)
            worksheet.write(apartment_global_index, 18, announce_price_per_m2)

            # write the CSV table file

        else:
            print('Still cannot extract information from ' + website)

        # update apartment index
        apartment_global_index += 1

workbook.close()