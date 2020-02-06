import win32gui, win32api, win32con, time, os
# import PIL.Image
from PIL import Image, ImageGrab

sensitivity = 0.17
catch_fish_click_size = 30
img_half_size = 40
total_count = 0
bait_time = 0

hwnd = win32gui.FindWindow('GxWindowClass', '魔兽世界')


def get_scale_regin():
    scale_x = 0.3
    scale_y = 0.6
    try:
        rec = win32gui.GetWindowRect(hwnd)
        rec_s = (int(rec[0] + rec[2] * (1 - scale_x) / 2),
                 int(rec[1] + rec[3] * (1 - scale_y) / 2),
                 int(rec[2] - (rec[2] - rec[0]) * (1 - scale_x) / 2),
                 int((rec[3] - rec[1]) / 2 + rec[1]))
    except:
        print("找不到窗口。")
        rec_s = False

    # regin = img =ImageGrab.grab(rec_s)
    # regin.show()
    return rec_s


def move_mouse(loc1, loc2, loc3, loc4):
    move_size = 20

    c_old = win32gui.GetCursorInfo()
    loc1_new = loc1
    loc2_new = loc2
    found_hook = False
    while not found_hook:
        while True:
            win32api.SetCursorPos((loc1_new, loc2_new))
            c_new = win32gui.GetCursorInfo()
            if c_new[1] != c_old[1]:
                # print("在", c_new[2], "位置处鼠标图标发生改变，成为", c_new[1])
                found_hook = True
                return c_new[2]
            time.sleep(0.02)
            loc1_new = loc1_new + move_size
            if loc1_new > loc3:
                loc1_new = loc1
                break
        loc2_new = loc2_new + move_size * 2
        if loc2_new > loc4:
            break


def find_fish():
    rec = get_scale_regin()
    # print(rec)
    win32api.SetCursorPos((rec[0], rec[1]))
    # time.sleep(1)
    return move_mouse(rec[0], rec[1], rec[2], rec[3])


# 将图片转化为RGB
def make_regalur_image(img, size=(64, 64)):
    gray_image = img.resize(size).convert('RGB')
    return gray_image


# 计算直方图
def hist_similar(lh, rh):
    assert len(lh) == len(rh)
    hist = sum(1 - (0 if l == r else float(abs(l - r)) / max(l, r)) for l, r in zip(lh, rh)) / len(lh)
    return hist


# 计算相似度
def calc_similar(li, ri):
    calc_sim = hist_similar(li.histogram(), ri.histogram())
    return calc_sim


def comp():
    image1 = Image.open('Image/ori.bmp')
    image1 = make_regalur_image(image1)
    image2 = Image.open('Image/comp.bmp')
    image2 = make_regalur_image(image2)
    result = calc_similar(image1, image2)
    # print("图片间的相似度为", result)
    return result


def grabe(pos, name):
    try:
        img = ImageGrab.grab(
            bbox=(pos[0] - img_half_size, pos[1] - img_half_size, pos[0] + img_half_size, pos[1] + img_half_size))
    except:
        print("没有捕捉到鱼漂。")
        return False
    else:
        img.save("Image/" + name + ".bmp")
        return True


def update_img(start_time, pos):
    global total_count
    comp_result = 0
    while time.time() - start_time < 30:
        got_fish = grabe(pos, "comp")
        if not got_fish:
            break
        new_result = comp()
        # time.sleep(0.2)
        if not comp_result:
            comp_result = new_result
        if comp_result - new_result > sensitivity:
            total_count += 1
            print("鱼鳔动了，开始拾取。总计第%s次。" % total_count)
            catch_fish_click(pos)
            time.sleep(3)
            break


def right_click():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
    time.sleep(0.2)


def catch_fish_click(pos):
    right_click()
    win32api.SetCursorPos((pos[0] - catch_fish_click_size, pos[1]))
    right_click()
    win32api.SetCursorPos((pos[0] - catch_fish_click_size,
                           pos[1] + catch_fish_click_size))
    right_click()


def mkdir(path):
    path = path.strip()
    path = path.rstrip("\\")

    # 判断路径是否存在
    # 存在     True
    # 不存在   False
    isExists = os.path.exists(path)

    # 判断结果
    if not isExists:
        # 如果不存在则创建目录1
        os.makedirs(path)
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        return False


def set_bait():
    global bait_time
    if time.time() - bait_time > 60:
        press_key(2)
        press_key(3)
        bait_time = time.time()


def press_key(cha):
    asc = int(ord(str(cha)))
    win32api.keybd_event(asc, 0, 0, 0)
    win32api.keybd_event(asc, 0, win32con.KEYEVENTF_KEYUP, 0)


def fishing():
    press_key(1)
    start_time = time.time()
    time.sleep(3)
    pos = find_fish()
    found_fish = grabe(pos, "ori")
    if found_fish:
        update_img(start_time, pos)


def shift_2():
    win32api.keybd_event(16, 0, 0, 0)
    press_key(2)
    win32api.keybd_event(16, 0, win32con.KEYEVENTF_KEYUP, 0)


def set_sen():
    global sensitivity
    new_sen = input("请输入灵敏度：")
    if new_sen:
        sensitivity = float(new_sen)


def set_foreground():
    # press_key("`")
    win32api.keybd_event(32, 0, 0, 0)  # enter
    win32api.keybd_event(32, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    if hwnd:
        win32gui.SetForegroundWindow(hwnd)


if __name__ == '__main__':
    # set_sen()
    mkdir("Imag  e")
    set_foreground()
    time.sleep(3)
    rec = get_scale_regin()
    while rec:
        shift_2()
        set_bait()
        try:
            win32api.SetCursorPos((rec[0], rec[1]))
        except:
            pass
        fishing()
        if win32gui.GetForegroundWindow() == hwnd:
            continue
        else:
            a = input("继续请按 1 :")
            if a == "1":
                set_foreground()
                time.sleep(3)
                rec = get_scale_regin()
            else:
                break
