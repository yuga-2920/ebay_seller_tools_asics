import eel
import sys
import desktop
import ebay_asics

app_name="web"
end_point="index.html"
size=(500, 700)

@ eel.expose
def main():
    ebay_asics.main()
    sys.exit(0)
    
desktop.start(app_name,end_point,size)
sys.exit(0)
