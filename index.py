import openpyxl as pxl
import random
import ast

# Open and setup doc
wb = pxl.load_workbook('../doc.xlsx', data_only=True)
sheets = wb.sheetnames

def read_file(filename, index):
    with open(filename, 'r') as file:
        lines = file.readlines() # Read the file
        if index < len(lines):
            return lines[index].strip()
        else:
            return None  # Index out of range

def read_excel(sheet, cell): # Read the value on a cell, in a specified sheet of the doc
    sheetName = str(sheet)
    ws = wb[sheetName]
    return ws[str(cell)].value

class Weapon:
    def __init__(self, name, hit, crit):
        self.name = name
        # self.rolls = int(rolls)
        self.hit = int(hit)
        self.crit = crit
    
    def display(self):
        return self.name, self.hit, self.crit

class Item:
    def __init__(self, name, defense, dodge):
        self.name = name
        self.defense = defense
        self.dodge = dodge
    
    def display(self):
        return self.name, self.defense, self.dodge

class Player:
    def __init__(self, level, role, weapon_ID, item_ID):
        self.role = str(role)
        self.level = str(level)
        self.weapon = Weapon( # Using the Weapon class in the Player class here
            name = read_excel("Weapons", str("C" + str(int(weapon_ID) + 1))),
            hit = read_excel("Weapons", str("D" + str(int(weapon_ID) + 1))),
            crit = read_excel("Weapons", str("E" + str(int(item_ID) + 1)))
        )
        self.item = Item(
            name = read_excel("Items", str("C" + str(int(item_ID) + 1))),
            defense = read_excel("Items", str("D" + str(int(item_ID) + 1))),
            dodge = read_excel("Items", str("E" + str(int(item_ID) + 1)))
        )

    def calculate_damage_rolls(self):
        damage = random.gauss(self.weapon.hit, self.weapon.hit/6)

        #for _ in range(self.weapon.rolls):
        #    roll_result = random.randint(1, 6)  # Roll a dice
        #    if roll_result >= self.weapon.hit_threshold:
        #        damage += 1
        
        if random.randint(1, 100) <= int(self.weapon.crit):
            return damage*2
        else:
            return damage

    def damaged(self, dmg):
        if random.randint(1, 100) >= int(self.item.dodge):
            return True, dmg - int(self.item.defense)/100*dmg
        return False, 0
        
    def display(self):
        return self.role, self.level, self.weapon.display(), self.item.display()

def main():
    # Values defined by a tk interface, written in a config.txt file
    config = ast.literal_eval(read_file("config.txt", 0)) # Convert the config file's info to array
    player_count = len(config)

    players = [] # Contains player objects

    for i in range(player_count):
        # Add a player
        players.append(
            Player(
            role = config[i][0],
            level = config[i][1],
            weapon_ID = config[i][2],
            item_ID = config[i][3]
            )
        )

    # Then temporary, just debugging for now
    for i in range(player_count):
        print(players[i].display())

        print("Player " + str(i+1) + " - Damage done: " + str(players[i].calculate_damage_rolls()))

        damage_taken = players[i].damaged(100)
        if damage_taken[0] == True:
            print("Player " + str(i+1) + " - Damage taken for 100: " + str(damage_taken[1]))
        else:
            print("Player " + str(i+1) + " - Dodged!")

wb.close() # Close the doc to avoid corruption

if __name__ == "__main__":
    main()
