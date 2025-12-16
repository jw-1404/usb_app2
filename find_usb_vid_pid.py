"""import usb.core
import usb.backend.libusb1
import usb.util

# 使用与主代码相同的后端
backend = usb.backend.libusb1.get_backend(find_library=lambda x: "libusb-1.0.dll")

if backend is None:
    print("错误：无法加载 libusb 后端。请检查 libusb-1.0.dll 文件。")
else:
    print("扫描所有 USB 设备...")
    devices = usb.core.find(backend=backend, find_all=True)
    if devices is None:
        print("未找到任何 USB 设备，或后端加载失败。")
    else:
        for dev in devices:
            try:
                vid = f"0x{dev.idVendor:04x}"
                pid = f"0x{dev.idProduct:04x}"
                dev_name = usb.util.get_string(dev, dev.iProduct) or "Unknown"
                print(f"VID: {vid}, PID: {pid}, 设备名称: {dev_name}")
            except Exception as e:
                print(f"读取设备信息时出错: {e}")
                """


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
                type_str = {usb.util.ENDPOINT_TYPE_CONTROL: 'Control',
                            usb.util.ENDPOINT_TYPE_ISOCHRONOUS: 'Isochronous',
                            usb.util.ENDPOINT_TYPE_BULK: 'Bulk',
                            usb.util.ENDPOINT_TYPE_INTERRUPT: 'Interrupt'}.get(ep_type, 'Unknown')
                print(f"  Endpoint Address: 0x{ep.bEndpointAddress:02x}, Type: {type_str}, Direction: {direction}")
    except usb.core.USBError as e:
        print(f"Error accessing device: {e}")