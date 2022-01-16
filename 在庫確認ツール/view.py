import eel
import sys
import desktop
import asics_inventory_check

app_name="web"
end_point="index.html"
size=(380, 300)

@ eel.expose
def main():
    asics_inventory_check.main()
    sys.exit(0)
    
desktop.start(app_name,end_point,size)
sys.exit(0)
