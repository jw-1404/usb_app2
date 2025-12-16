import usb.core
import usb.backend.libusb1
import usb.util

# 使用与主代码相同的后端
backend = usb.backend.libusb1.get_backend(find_library=lambda x: "libusb-1.0.dll")
dev = usb.core.find(backend=backend, idVendor=0x04b4, idProduct=0x00f1)
if dev is None:
    print("Device not found. Ensure it is connected and VID/PID are correct.")
else:
    try:
        dev.set_configuration()
        cfg = dev.get_active_configuration()
        print(f"Configuration: {cfg.bConfigurationValue}")
        for intf in cfg:
            print(f"Interface {intf.bInterfaceNumber}, Alternate Setting {intf.bAlternateSetting}")
            for ep in intf:
                direction = 'IN' if usb.util.endpoint_direction(ep.bEndpointAddress) == usb.util.ENDPOINT_IN else 'OUT'
                ep_type = usb.util.endpoint_type(ep.bmAttributes)
                type_str = {
                    usb.util.ENDPOINT_TYPE_CTRL: 'Control',
                    usb.util.ENDPOINT_TYPE_ISOCHRONOUS: 'Isochronous',
                    usb.util.ENDPOINT_TYPE_BULK: 'Bulk',
                    usb.util.ENDPOINT_TYPE_INTR: 'Interrupt'
                }.get(ep_type, 'Unknown')
                print(f"  Endpoint Address: 0x{ep.bEndpointAddress:02x}, Type: {type_str}, Direction: {direction}")
    except usb.core.USBError as e:
        print(f"Error accessing device: {e}")