import mouse_action
import price_excel
import concurrent.futures

if __name__ == "__main__":
    price_excel.byd_import()




def bydupload(app, action_name):

    executor = concurrent.futures.ThreadPoolExecutor(max_workers=2)
    executor.submit(mouse_action.ok_btn_action, action_name)
    executor.submit(app.Application.Run(action_name))

    executor.shutdown()
    #print('並行実行終了')

   