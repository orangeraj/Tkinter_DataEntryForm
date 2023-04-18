import qrcode

# Generate QR code for a website URL
url = "https://www.google.com/maps/place/HANUMAN+Mandir+NAVADE+Taloja/@19.0522203,73.1001132,17z/data=!4m15!1m8!3m7!1s0x3be7e9f07164197b:0xcf882f74c583ad1a!2sNavde,+Taloja,+Navi+Mumbai,+Maharashtra!3b1!8m2!3d19.0511354!4d73.0966019!16s%2Fg%2F11rdhpv9j!3m5!1s0x3be7e9faa0718f6b:0x4b920f0572329ea2!8m2!3d19.0498563!4d73.099515!16s%2Fg%2F1hm33p4fr"
img = qrcode.make(url)

# Save the image file
img.save("example.png")
