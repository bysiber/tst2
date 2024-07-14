from _utils.sharepoint import SharePoint
from _utils.common_utils import send_email
from str_distribution.config import Paths, Emails, TRANSFER_LOGIC_URL
import traceback
from _utils.logger import logger
import  config as config
from pathlib import Path
import base64
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta

def generate_paths(vpo_data: dict, processed_directory: str, report_name: str):
    if processed_directory and report_name:
        vpo_data['fpaths'] = (Path(processed_directory, report_name))

def send_report_emails(emailing_dataset: list, start_date: str, end_date: str) -> None:
    start_date_obj = datetime.strptime(start_date,"%m.%d.%y")
    next_day_date_obj = start_date_obj + relativedelta(months=1)
    next_day_date_formatted = next_day_date_obj.strftime("%m.%d.%y")

    failed_datasets = []
    for dataset in emailing_dataset:
        # Prepare Email
        email_subject = f"{start_date} to {end_date} - BI-WEEKLY Pcard Outstanding"
        email_body = f"""
                <div style="font-family:Calibri, Helvetica, sans-serif;">
                    <p>Dear VP of Finance and Property GMs,</p>
                    <p>Please see the attached Bi-Weekly Pcard Compliance report for {start_date} to {end_date}.  This report summarizes the Pcard bi-weekly spend for your properties and shows transactions that still need to be coded, have receipts uploaded, and the sign-off completed (highlighted on the report). Per policy, these actions are due as soon as the transactions are posted in Works.</p>
                    <p>The first tab contains pivot tables summarizing the spending for your properties and the 2nd tab contains all of the transactional information. We hope you will contact your properties and ensure they are coding the transactions, uploading the receipts, and signing off so that we have all transactions coded by month's end. Pcard limits may be taken to $1.00 on {next_day_date_formatted} for those who have not coded their transactions for {calendar.month_name[int(start_date[:2])]}.</p>
                    <p>This report was pulled today {datetime.now().strftime("%m.%d.%y")}, if the coding has been completed, the receipts have been uploaded and sign-offs are done, please disregard.</p>
                    <p>Please let me know if you have any questions. Thanks.</p>
                    <p>EMAIL: <a href="mailto:PCARD@HIGHGATE.COM">PCARD@HIGHGATE.COM</a><br></p>
                </div>
                <table style="width: 471pt; border-collapse: collapse; border-spacing: 0px; box-sizing: border-box;">
                    <tbody>
                        <tr>
                            <td style="padding: 0in 11.25pt 0in 0in; width: 67.75pt;">
                                <p align="center" style="line-height: 12pt; margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif; color: blue !important;">
                                        <a
                                            href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.highgate.com%2F&amp;data=05%7C02%7CRGioia%40blueoceansp.ai%7C97e340e09ca242351d6508dc5308d3b4%7Ceb8172eda7a8432cb2620327e4eb2097%7C0%7C0%7C638476544592277339%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&amp;sdata=8wiKkcob3NbPbz%2BCnMZ7jQywv8ssNI4dNRSDHnu4BQM%3D&amp;reserved=0"
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            data-auth="Verified"
                                            originalsrc="https://www.highgate.com/"
                                            shash="QI62XVEPKAdFZpNbtFY4UjbHykAS46sHsk3uIzhD3PmmGWd8AYzUSy7yonmr/tmAkIEVMjn6NNR+/19pbWTn9PMIUiO8geHXwksY6UhlKZ5dZoEk9s7gb9zH+P8JGq5KpA3TaksfKwb0aGFI7iCkUMuXe3D/D+tudDKnslV3zZQ="
                                            title="Dirección URL original: https://www.highgate.com/. Haga clic o pulse si confía en este vínculo."
                                            id="OWAf9708882-463d-4f12-5b34-a9fb65f6ad50"
                                            class="x_OWAAutoLink"
                                            data-loopstyle="linkonly"
                                            style="color: blue !important; margin-top: 0px; margin-bottom: 0px;"
                                            data-linkindex="16"
                                        ></a>
                                    </span>
                                </p>
                                <div>
                                    <a
                                        href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.highgate.com%2F&amp;data=05%7C02%7CRGioia%40blueoceansp.ai%7C97e340e09ca242351d6508dc5308d3b4%7Ceb8172eda7a8432cb2620327e4eb2097%7C0%7C0%7C638476544592277339%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&amp;sdata=8wiKkcob3NbPbz%2BCnMZ7jQywv8ssNI4dNRSDHnu4BQM%3D&amp;reserved=0"
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        data-auth="Verified"
                                        originalsrc="https://www.highgate.com/"
                                        shash="QI62XVEPKAdFZpNbtFY4UjbHykAS46sHsk3uIzhD3PmmGWd8AYzUSy7yonmr/tmAkIEVMjn6NNR+/19pbWTn9PMIUiO8geHXwksY6UhlKZ5dZoEk9s7gb9zH+P8JGq5KpA3TaksfKwb0aGFI7iCkUMuXe3D/D+tudDKnslV3zZQ="
                                        title="Dirección URL original: https://www.highgate.com/. Haga clic o pulse si confía en este vínculo."
                                        id="OWAf9708882-463d-4f12-5b34-a9fb65f6ad50"
                                        class="x_OWAAutoLink"
                                        data-loopstyle="linkonly"
                                        style="color: blue !important; margin-top: 0px; margin-bottom: 0px;"
                                        data-linkindex="17"
                                    >
                                        <img
                                            data-imagetype="AttachmentByCid"
                                            data-custom="AAMkAGE0NDhmNmY4LWI5OTQtNGYxNy05MmFiLTE1YmY1MzNlNDBhNQBGAAAAAADMrIwFplGzS7yzyq7lAXxaBwAZmXyA7SUvQ72Ye0QZ%2FmlMAAAAAAEMAAAZmXyA7SUvQ72Ye0QZ%2FmlMAACKdmpdAAABEgAQAKiWDGowO39Om6ghceLYjrg%3D"
                                            naturalheight="0"
                                            naturalwidth="0"
                                            src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPAAAADwCAYAAAA+VemSAAAAAXNSR0IArs4c6QAAIABJREFUeF7tXQeYFMXzfY0CBgRBTCgqURBQkKCA5HRHkHgEyQgoUUAJguQgSQxkkJxzBiUnAQ8QJf4QUEkiiIoCgqD0f9+we//dvdnbMDObruv79jtlZ6q7a/pt91RXvRJQoiygLBCxFhAR2/Nk1HEp5ZMAnrN/ngDA/38MQHoAD9v/PgggDYCUAB4AkNrNRP8A+BvAbQDXAFwH8AeAK/a/lwD8Yv/8COAnIcSFZGTmiByqAnCYPDYpJcGXB8CLAPIByAYgqx2094WomzcJZAA/ADgF4BCAgwCOCCH4I6AkxBZQAA7BA5BScrUsDKAEgPwAXgKQBUCKEHQnkCbvAOAq/R2AbwHsALBXCMFVXUkQLaAAHARjSykfBVAaQDEAxQEUAHBvEJoOZhP/2sG8E8Au23i3CiF+DWYHkmNbCsAWPHUp5T0AigCIBRADoKAZq6uUEn///Tdu3ryJGzfu/r158x/t7z///IM7d+5Ayju4dYuvucDt27dw545EqlSpIMTdR50y5b1ImTKV9pf/nipVajzwwAO47777cP/992v/fe+9pvy2cJXeD+ALAOsAxAsh/rPA3MlapQKwSY9fSnm/HbBxthWoIoAMgai+ffs2Tpw4gePH/4ezZ8/aP2dw5sxZXLjwM/79lwudtZIxY0Y8+WQmZMqUCc8++yyyZcuGrFmzImvWbHjsMfrOApLfbTuQ9QAW237c1gohbgSkRd3kYgEFYAMTQkpJTy9X2Hq21aaa3Qvss8Zr165i//79OHLkCI4dO6Z9Tp06GRSQ+txJtwvTp0+PF154Ablzv4A8efIif/78GsAdK7yPeukAWwVgAVdoIQQ95EoCsIACcABGk1IWBdDK5pGtBSCdryrOnTuHffv2Ij4+Xvt7/Phxbdsb6ZI2bVrkz18AhQoVQvHixVGgwMtImZKnWT7JXzbP+1IAk4QQu326Q12UYAEFYB8ng5SSW+ImAFraj3u83sl30927d2HLli3YvHkTTp8+7fWeaLiA79MEc4kSJVGmTFnkzp3b12EdATDF5pGfIYTglluJFwsoAHsxkJSSzqhO9tXWPTgi0d0XL17E2rVrsHHjRnz99R7NuZTc5cknn0TZsuUQExOD114r4YuTjEbjqvyJECI+udsvqfErAOtYR0rJ89iqtoil9+xntUnOocuXL2PdunVYuXKFBlp6i5XoWyBdunQoX74CXn+9OkqWLOkLmHnGPBLAaiFE5L9vmDwxFICdDGp3SjUF0AXA80nZ+tatW9pKu2DBfOzatSsq3mVNnlte1T3yyCOoXr0GateujRdfZCxLknIcwChbdNpMIQQjxJQAUAAGuGKmAtAcwAcAnk5qZnz//XHMnTsXixcvwp9//qkmkUkWoGe7YcNGqFWrFtKkeSgpredsx9mDAUwVQtwyqfmIVZOsASylZMRCYwB97DHHug+SZ6+rV6/C9OnTtGMfJdZZgMEkNWvWQsuWLZEjR86kGmKM9kD7imz94bh1QzakOVkCWErJcde37UD6A8jhyYLXrl3DvHlzMWXK5zh//rwhQ6ub/bdAyZKl0Lp1a/BvEufMJ2ybqL4A5gshkp3zIdkBWErJOORPk3JOXbp0CZMnT8KcObNx9epV/2eeusNUC+TJkwcdOryD2NhYpEjhMd+DMdgdhRAHTG08zJUlGwDbc2oH2dLjmnmKS/79998xZsxnmDlzpjr+CcOJmz17DnTu3BlVq1bzBGR6qafTl5FccpmjHsD299yO9u0yc24TyV9//YWxY8do77hMFlAS3hagw+u997qhQoUKHt9+7Nvqz4QQUf1+HNUAllLybOJz2xliIb0nzSALbpXHjRurtsrhjVnd3hUuXBh9+vTT4rE9yD5GzgkhmLcclRKVALZnBvW20cZ09ZR3u27dWgwcOBBnz56JygebnAZVu3YddO/eA4z40hGuwCPosY7GDKioA7CUkknz0wDonkEw46dv397YvVvFzUcTyHn81LlzF7Rq1dpTdNf3POsXQpBsIGokagBsf9flqttTb9Xlu+2IEcMxdeoUFTUVNdM38UB4djx06FAUKfKKp9V4iH01jop346gAsJSS5G+z7JQ1iR7ctm1b0aNHD5w7dzbipy7DD59+OjMyZnwEDz+cHhkyZABzdB9++GGNVYMfMm1Q0qa9m+lIZo4bN+7mz1+/fl1j8OAP2pUrV/DHH39onytX/sCFCxe0D6+JdGncuDF69uzlKaqL269GQgiS9UW0RDyApZRM8RvN+er+JDgx+/fvhyVLSAIROfLQQw9pCfPPP58TOXM+j+eee04DbebMmZE6tdeEKMMDZQDLmTNncPr0T1oKJBlCTpz4Ht9//31EgZuMIiNGjNQCQXSEecgdhBAzDRsshAoiFsBSSlKtjref6yYy4aZNm/Duu53x22+/hdC83pvmu9tLL+XX8mcLFiyEvHnz4oknSP0cfsIsKzr9Dh48hIMHv8O33x7AgQMHtBU9nKVx4ybo06evtjvREZ4btxNCROT5YUQCWEpJzmQuq4nOD3g0NHDgAMyYwecSfsIVlO9nZcqUwauvFtWS3U0ikQvJYBknfvDgQezdG48dO3Zgz57dYRkEwyCQMWPGglFdOkJq3DpCCHJfR5REHICllJVsHua5eqRxR48eRbt2bXHyJMNjw0eeeeZZVKxYUQMtwethJQifDhvoCX9A4+O/1ggN1q//EqQRChchzc8HH/RGixZv6nWJVSoaCCG+DJf++tKPiAKwlJJ5usNtXkTStroIV1y+75LVMRyEoK1atar2yZePxRaSp5Cwb82a1VixYgXOnAkPSiE+kxEjPkKaNIkC80h7200IwbzjiJCIALCdZ/kzAG3drcr3rx49uoeFo4rOpxo1aqBevfrae60SVwt888032nNatmxpyCPfSJP7+edTPKUsjrMnRoQ9j3XYA9heM2i+rXRHFXdA8Be9VauW4NY5lFKsWDHUr98AsbGVo3p7bJaN+aNLNpOZM2eENL+aK/Bnn43xFFO9himn4V4DKqwBLKXMZOcPftl98mzdugXt27cLGSsGz1pr1aqtJZ4//3wus+Z2stNz6NAhTJs2VVuVg0Fa725g5hl369Yd7dq118s5ZmpiVSHEz+H6YMIWwFJKclLRofCsu/FmzZqJDz7oFZKIKgZN0AnCowkGVSgxxwK//PKLllgye/askGSExcXVxbBhw/X4rBksX1EIQU6usJOwBLA96Z41dVzqePAccvDgQZg4cULQDUngvvXW22jWrDkefJDFBZVYYQFGh40fPy4kqZ0M+Jg4caJe9BZrJ8cKIb6xYsxGdIYdgO1VD1gMy6XiAd+b3nmno/buFExRwA2mtf+/rbvkCqO17XUwt9bMNZ4zZx5YH8pNyGBIEIdVFkxYAVhKyXq5RKgLLSFpbRo3bhhUhwffcd98syU6duzojSUxNDM8mbRKR+WQIYOxZk3wfrhZ0G3BgkV46qmn3K1MfqUqQghyVYeFhA2ApZSsm8vqdQ84W4ahkPXq1dWq9QVLqlWrhvff74nMmZ8JVpOqHS8WYIRXr169QFrfYAjDWQliHje5CUMu+U78VTD64a2NsACwlJL1cze5b5uZGfPGG/Vx8uRJb+Mw5Xumog0bNgyFC7OaipJwswC30kwH/eijkUFxdDHTiyDOlSvRKQO30+WFEGT8CKmEHMBSynx28LKKfYIwC6ZhwwZBKQjG7fI773RCmzZt/amqF9IHl5wbZ0JF167v4auvrF8EedKwfPlKLSPMTX4FUE4IcSiUzyKkAJZSZrcZge8TLuk3BG+dOrXAowWrpWjRohg6dLjeVsnqppV+AxbgicScOXMwcGB/y1djbqcXLVqiB2JO0BJCiOBsEXXsFTIASylZwoT0Jpmd+8Vtc1xcbctXXq66PXq8j5YtW/lbnNrAtFO3mm2BH374AR07dsB33zGhyDoh39aKFav0eLeYrVFUCBGSrI2QAFhKSUcV9z8uAcN0WHHltfqdlyl8DKHTebexbgYozZZZgO/Gw4YNxYQJTA+3TrJly4YlS5bpBfActIM46DnFQQewPTGBtV9fdzY1j4pq1qxhubeZR0OkWnHQzlj3uJXmYFtgw4b16NTpHZDn2yphZtmiRYv1gnlWsoa0ECKoCRChADAzPdo4G5hBGvXr17X0nJfRUyNHjtLS+5RErwV4bty8eXNLj5vKli2LqVOn4557EmW1jhdCJMqYs9LaQQWwlJIFs8nRmyB0Rrz99luWRljxeGjSpEkgK4OS6LcAOb34XswV2SphLPyQIR/qqWc+scsct6oP1Bs0ANuZNBhO4/KzNWjQQEtjmytWrITPPhut4petnEVhqJsLA+fWpEkTLesdq0K0atXKXT+30IzWCgqzR1AAbD8uirdVBUzvPFpmFfXs+b5lBmbyAd93k6hoZ1nbSnF4WICBH2RquXOHdc/MFc6r2bPnokQJRgC7COl5igTjeMlyAHvyOJOruUmTxpYYlu8mgwYNQaNGjcx9YkpbRFqAcdTt27e1JCkiXbp0WLt2HUih5Casx1TMarbLYAB4KktaOA+OjobKlWMtScYnYRypUkqVKh2Rk0112hoLbNmyGa1bt7KEApeEDqtWrQYpgt1kuhDCZe6bPTpLAWwnXZ/h3Gl6nKtXr2YJDU6aNA9ptX1VLLPZ0yQ69DH0slmzJpaAOC4uDqNGfaJnKNZjsozj2DIA2xk1GOztQv3HczorKiUw8Hz27DnJmgEyOmBm7SisBPGoUR+DzB5uwjo1Ba1i9LAEwFJKFufZA6CA82BI/UoqHLPl0Ucfw/z587UyJEqUBbxZwCoQcwvN92Gd40pya70qhLjlrW/+fm8VgHlA1sO5M2SOrFq1sum8zUmkfPlrC3V9MrIAz4hbtnzTdCcqKz+sWrVGL6ttmBDCBRNmmNt0AEspXwOw1fm8l2z9MTGVTK+YwHfexYuXeCqXYYZ9lI4otsCiRYvQpUsn00fYtm07jRDCTXg+XFoIsdPMBk0FsL3gGAO7XUKeuG02u1YRwTt37jwUKOCySzfTNkpXMrDA6NGfYfjwYaaOlFS1XFh0ahSz5s+LtpXYtGpwZgM40daZVQLp+TNTWAxs+vQZ6qjITKMmY13vvfcuFixg7QDzhAQA69dv1DtaGiqEMC16yTQASylfspFg0+t8r8MMrM9bpkwp00t8jhz5kVa+RImygBkWYD0tsr/s3m0u4WSrVq21sqZu8q+NuLGwEMKUBGZTAGxPEaTXuZBzZ604MurQoaPGpK9EWcBMCzAXvUqVWJw/f940tQy1JB2PzmvefgCvmJF6aBaA37Wd9450Hvn27dvQsOEbphmDiqpXr4HRo8coBg1TraqUOSxw6NBBLSedTlezhOQRa9d+oVcDuqsQwgUzgbRpGMBSSpLnkvM1IWDj77//RrlyZXHu3NlA+qR7Dw1BShOdcDXT2lCKlAXmzZuLbt26mmqIfv36axzjbnINQC4hhKEl3wwAzwHgstQy++PzzyebZoQkAsZNa0MpUhZwWIC5xCy2ZpawCuK2bTvw2GMulYKofq4QoqGRdgwB2E7GTlbJBD3Hjh1DTExF0w7I+R5Bj3OZMmWNjFPdqyzgswWuX7+OypVjQMI8s6R27Tr45JNP3dVJO6tlwPy4AQNYSsl76bhyYUGvW7eOqd485bQyawopPf5YgCyXNWpUNzUFcc2atXjxRR7WuAjz5BlmSTD7LUYAzHOcec4trlu3VkvZMkvovVu6dLmeA8CsJpQeZQGPFvj0008wcqR57DjMklu6dJleew2FEHMDeRQBAVhKmRLAMQDZHI3Sc1emTGmQNd8MIQndl19uAAtNKVEWCIUF/vvvP1Sv/rqpnNMTJ05C5cpV3IdzCkBuIcRtf8cZKIDJKkl2yQRhKUhy85olHlKzzFKv9CgL+GQBFtWLjY0xLQknS5Ys2Lx5q96usq0Qwm9ia78BLKUk7QBLSWRyWIA8vK++WgTkdjZDypcvj2nTXHgAzFBrqo6BAwcYZtJkQWlWhY9kadu2DQ4cMFb3mitS7959wtYMo0Z9hI8/HmVa/8hmSVZLN7nAHa0Q4oY/DQUC4I42cjoXd9qHHw7BuHFj/WnX47V0uW/atAWZMiX8Ppii12wlzGJhNosRIWPmlClkHIpciYurA5b+NCJJsFkYUWvavQy1rFChHE6d4k7XuDB/fceOnXpMqZ2EEIlc1Um16BeA9VZfVlIvWvQV0wpMDRo0GE2bNjNuJYs1KADfNXByADDHuXPnTjRoUM+0WfXBB71B1lQ3+RlAdn9WYX8BnOjdd8CA/pg8eZIpAytUqJBWeyYSaGAVgJMXgDlavi6sWsUKKoFL2rRp0bHjO2jevIWn8j7thBAu/iVTVmApJbOMGDKZ4Hm+dOkSihV71ZTYUYL2iy/WgyGTkSAKwMkPwKycWaJE8YDme8qUKdGsWXOtDjUjC5OQ0/ZVmFlLXsXnFVhKSbauBc4aBw8eZFpFuIYNG2HoUHMTq72O3sAFCsDJD8Ac8YgRw/HZZ369pqJmzVro1q0bnn7apZJuUrOvnhBioS/T0x8A01PxqkMp688UKVLIFM8z2TX4Up8xY0Zf+hwW1ygAJ08AM8yyRInX8Ouvl7zOw+LFi6NXrw8CYUrdYyvNUtRrA77WRpJSErgursbJkydjwIB+vrTh9RoPL/Re7wvlBQrAyRPAHPXcuXPQvXs3j9OPRO8s6cMqhgaERcMZqpyk+LQCSykZ5tXAoYkFlV97rZgpyc88Ltq5c5cei5+3vof0ewXg5Atgzn8eK7kXon/88cfRtWs3jRvaBEfsPCGE14R6rwCWUj4CgDmLqR2IWbFiOdq3b2cKgIYMGYrGjRuboiuYShSAky+AOXLS0rZocbdqCmMXyERJCh2W9jFJyCrwlBDit6T0+QJg8m5+7KykRo3XTSnGzTjnrVu3R2SyggJw8gYwR8+i9Kw93alTZzzyCNc506WLEMIFe+4t+ALgwwDyOG48ceJ7lC1bxpSeRnK8swKwAjCTHVgJ00I5KoRIwJ5eO0kC2J6w70JE3a9fX0yZ8rnhPtOlTs8zKWIjURSAFYCDNG9fE0J4TPj3BuApAFo4Onrr1i0ULFgAV65cMdz3/v0HoEWLNw3rCZUCBWAF4CDNvWlCiAQM+ryFllLSacXDrrSOm5YvX4YOHdob7jfDyeLj9+kFcxvWHSwFCsAKwEGaa3/ZUncfs5HB61JlelyBpZQ1ALjQBzCYm0HdRqVNm7baOVkkiwKwAnAQ529NIcRyv96B3c9+L1++rG2f79y5Y6jffOfdvftrPPHEE4b0hPpmBWAF4CDOQY9nwrorsJTyAQAXnbmeZ82ahZ49jVdHjIYcWD44BWAF4CACmBzSjwsh/vbpHVhKWcd2dOSSrW4W2ySZNsi4EemiAKwAHOQ5XEcIscRXADPriNlHmly8eBGFCxeElAExXya0yW3znj3xVp+dBcWuCsAKwEGZaP/fyAIhRKKKfom20HbGyV9t0VcJSYvTpk1Fnz69DfeXicyMFY0GUQBWAA7yPP4TwKPuzJV6AGaY1WbnzjVu3Ahbt24x3F+Wl8iaNathPeGgQAFYATgE87CsEMIFiHoAJk1iQnWnmzdvIm/eFwJiIXAeYL58+bQqbdEiCsAKwCGYyyNsjiyXLawegFl4OKH+w5Ytm9GkifFsoR493ke7dsaDQEJgNN0mFYAVgEMwF7+znQfnd27XBcBSyscBkJ824d/57st3YKPCnN9oqrKgAKwAbBQTAdxPL/KTQgge8WriDmDyZs53VszE/dOnybMVuOTJk0cjrIsmUQBWAA7RfK4vhEjgpnMHMNnZ2zo6du7cOY3z2ahEY4VBBWAFYKO4CPD+8UKIBIy6A/gQgLwOxWYlLyxevASvvJLAhxdgv8PrNgVgBeAQzcjDtoCOfIm20FJKnvv+DiCF48uePd/HrFkzDfWTjJOHDh2O2LxfT4NXAFYANgSMwG9mMkIGIQTPhf//HVhKWQGAy4tqxYrlcewYq4gGLjExsZg82TgBQOA9sOZOBWAFYGtmlk9aKwohNrgDmOXh+jtuv3btKvLkecFw9lGkktZ5M6MCsAKwtzli4fd9hRAD3AG8FEBNR6Pbtm1Fo0YNDfdh/fqNEVMuxZ/BKgArAPszX0y+dpkQopY7gE+wJoujIZYLZdlQI0K6zSNHjpnBkWukG5bcqwCsAGzJxPJN6UkhRI4EAEspH7SluJK6I8GBReoceqGNCAtYz5lDTvjoEwVgBeAQzmoGdKQTQlzVjpGklKzDssu5Q+XLl8Px4yxGGLh06fIuOnfuEriCML5TAVgBOMTTs5gQYrcDwK1tIVoTHR1iRfKcObODJSSMyPTpM1GuXDkjKsL2XgVgBeAQT863hBCTHAAeDSAh0+Do0aOoVImnSsaEyftPPfWUMSVhercCsAJwiKfmGFvZlQ4OAK8FEOvokBkRWAzgOHbM2BY8xAZKsnkFYAXgEM/P9bYtdCUHgI/ainfndnSIBYxZyNiIFC5cGEuX6jJhGlEbNvcqACsAh3gy/iiEyOoAMNnu7nd0qGvX9zB//jxD/WPFQQZxRKsoACsAh3hu00H1gJBSZrKXD03oD6uuffWVx3IsPvW7T5++WrnFaBUFYAXgMJjbWQjgYgBc0FqsWFGcPXvGUP8mTZqM2NjKhnSE880KwArAYTA/SxLArAI+x9EZUsdmzfqc4SMk8l+RBytaxQwA582bN6ILvPHZjh07BqdOnTL0mF96KT9q1GAlHyVJWaBcufLIkiWL8yUNCWCXAt7Xr19Hrlw5DVvy0KEjePjhhw3rCVcFZgA4XMem+hWeFpgwYRKqVKni3LnOBPBAAB84/vW3335D/vwvGhoBY6CPHTtuSEe436wAHO5PKPr6pwPggQTweABvO4Z77txZFC1qjD2D3M/kgI5mUQCO5qcbnmPTAfAEApg1kFgLSZOTJ0+iTJlShkbw8ssvY8WKVYZ0hPvNCsDh/oSir386AF5EAG8CUNYx3MOHDyM2tpKh0TP+mXHQ0SwKwNH8dMNzbDoA3kwAuxC579u3DzVrVjc0gjp16uDjjz81pCPcb1YADvcnFH390wHwtwTwTwCedQyXARwM5DAiDOBgIEc0iwJwND/d8BybDoB/IoB/Jtu7o8s7duzAG28kqmLo14iYA8xc4GgWBeBofrrhOTYdAP9CAF9xLiW6ffs2NGzI2I7ARQE4cNupO5UFPFlAB8B/EsCk0nnIcdPGjRvRvHlTQ1bs06cfWrVqZUhHuN+sVuBwf0LR1z8dAF8ngMmvkyBmALhv335o2VIBOPqmkBpRKC2gA2AkAvC6dWvRurUx8A0Z8iEaN24SyrFa3rZagS03sWrAzQIKwCZOCQVgE42pVPlkAZ8ArLbQPtkSCsC+2UldZZ4FdAB8zRInlnoHNu+hKU3KAg4LePJCq2OkAOaIWoEDMJq6xZAFdAB8kSvwBQBPODSrQA7fbGwGgMnc+cwzmX1rMEyv+umnn/D336RUC1zSpUsXtfTDDqv873//M1wokFU+We3TSc5YEkrZuvVb6N2bxQ6jV8wAcMWKlTBlytSINlJcXB3s2bPb0Bji4uIwatQnhnSE+83PPpvZMICXLFmKIkVecR7qD5YkMySHB6IAfHceKQB7/+lgqd7cuXN5v9DLFatXrwHph5xES2ZwSSc8dOgQKleOMdRY+fLlMW3aDEM6wv1mBWAFYF/n6Jkzp1G8OLkjjcmGDZuQK5fLD4GWTuiW0H8CZcqUNtRSwYIFsXz5SkM6wv1mBWAFYF/n6KFDB1G5ssu7q6+3uly3fftOd1K7xZZQ6mTLlg1bt24PqJORcpMCsAKwr3N1x47teOONBr5e7vG6r7/ei0yZSOOeIBqljumkdg899BCOHo3eukg0nwKwArCviFy0aCG6dOns6+Uer/vuu0PIkCGD8/caqZ0ltLJHjhxD2rRpDXc6XBUoACsA+zo3P/nkY3z00UhfL/d43alTPyJVqlTO32u0spYQu3/xxXrkyZPHcKfDVYECsAKwr3PTjFpjHqiaNWJ3S0qrTJ06DRUqVPR1jBF3nQKwArCvk5bvv3wPNiI5cuTE5s1b3FVopVUSFTerVy8Ou3btMtIe+vcfEPFlQ5IygAKwArCvAClZ8jX8+OOPvl6ue13p0mUwa9Zs9++yEsAsMXrd7PKizZu3wIAB9I9FpygAKwD7MrNv376NHDmy4b///vPlco/XNGzYCEOHDnP+ngrvt6zA96uvFsWiRYsNdTqcb1YAVgD2ZX4eO3YMFSuW9+XSJK95//2eaNu2nfM1LgW+1wJIOGletmwpOnbsYKhRFjZjgbNoFQVgBWBf5vbKlSvQrl1bXy5N8hodn9J6IUQlxwo8GkB7h4ajR4+iUqUKhhvdv/8AHnvsMcN6wlGBArACsC/zcuTIEfj0U+OJGl99tQvPPJNA386mxwghOjgA3NrGDT3R0SHu23PmzG64RvDcufNQokRJX8YZcdcoACsA+zJpW7Rohg0bNvhyqcdr7rvvPhw/fgIpUqRwvuYtIcQkB4CLAnBxO5cvXw7HjxuLpurevQfatze2FTc0cgtvVgBWAPZlehUokB+XL//qy6Uer2Eh+HXrvnT/vphtC73bAeA0AP4EkADx9u3bYcWK5YYarlChAqZOnW5IR7jerACsAOxtbppRqpdt1K5dB5984lJrjFTQ6YQQVzUAU6SUJwBkd/z/uHFj8eGHQ7z1Mcnv06dPj4MHDxvSEa43KwArAHubm6tWrUTbtm28Xeb1ex2OuZNCiBy80RnASwHUdGjbtm0rGjVq6FW5twuYlcTspGgTBWAFYG9zesCAfpg8ebK3y7x+z9Rcpug6yTIhRC13AJMDp7/jIrII5MnzgmEakI8+GoW6det57WSkXaAArADsbc6SGIMEGUbk3nvv1RxYbkkMfYUt1a7wAAAYzklEQVQQA9wBzHOj9c6N8QCaB9FGpFat2vj008+MqAjLexWAFYCTmph//PEHXnopH19NDc1fUuiQSsdNKgohNNe28xY6HYDfnR1ZPXu+j1mzZhrqwCOPPIIDB76DEAlNGdIXLjcrACsAJzUXV69ejTZt3jI8XZs1a46BAwc56+EvQnohBJ3O/w9g/o+Ukut9XsfVy5cvQ4cOCfEdAXdm7dovkC9fvoDvD8cbFYAVgJOal927d8PcuXMMT91x48ajWrXXnfUcEUIkYNRlWZRSjgOQ4DY7d+4cihZ1obEMqEPReB6sAKwAnBQYihZ9FTxGMiLctXL3yl2sk4wXQiTEZroDmN6m+c5Xv/ZaMZw+fdpIP1C4cGEsXWrsTNlQByy4WQFYAdjTtDpy5AhiYoznwpMQg8QYblJfCLHA8W/uAH4cACs1JPx7nz69MW2aMfJx/pLs3bsfjz9O9dEhCsAKwJ5mslnxz2+/3Qa9en3g3Azff58UQlzUBTD/UUr5LYCXHBds2bIZTZo0Now65gYzRzhaRAFYAdjTXDYjDJm6dXIJvhNCuDC7J3INSymH2xL8uzo6d+PGDeTN+wJu3bplCHvRlh+sAKwArAcI1ooqUaK4Iazw5tSpU+Pw4aNgIoOTjBBCdHP+Bz0AlwGw2fmihg3fwPbt2wx1ipkUTC/MmDGjIT3hcrMCsAKw3lw0i4GSfHLMAXaTskIIF2IsPQCnBMD0CZ4La8J3YL4LG5VoqhusAKwA7I4HBm2Q/4qrsFFh8gKTGJyE576PCiFuJ7kC80spJb1cdR0XXrx4EYULFzQcVfL887mwcSNLMUW+KAArALvP4r1741GrVkI6QcCTnOGTPD4iq42TLBRCJIpJ1g2PklIS+qyZlCC1a9dCfPzXAXfKcePKlatRoEABw3pCrUABWAHYfQ5269YV8+bNNTw1SYJBB5abxNkCOBKRzHkC8AO2gA66qpknrMmsWbPQs2cPw51r0OANDB8+wrCeUCtQAFYAdp6D169fR6FCL+PatWuGp+aQIUPRuLHLyQ+VPm5zYCWqpO4xQFlKyZ+ShIpMly9fRsGCBQxnJz344IOaM4t/I1kUgBWAnefvzJkz0KtXT8NTmttn4sOtBtI8IQQrqCSSpABcw3YevMz5jgYN6mHnzp2GOxkNpO8KwArADiDQeVW2bGmcPHnSMDZiYmIxefLn7npq2s5/dUMZkwJwagCXACRUKDMrueHppzNj586vcM899xgecKgUKAArADvmHo9YedRqhujQx/5l3z7f9GsF5sVSSsZQNnfc+M8//2jb6D//1DKZDMmECZNQpUoVQzpCebMCsAKwY/41a9YEmzYZP11hjARDjrmNdpIZQohmnuZ6kkm6UkqGlLjsmfv164spUxIt8X5jqVChQli2bIXf94XLDQrACsC0gFmVF6irVatW6NOnn/sUf00I8VVAALavwiyv8IJDAalmGetphixcuBhFi5LRNvJEAVgBmBYgaR3J68yQLVu2Int2javOIUeFEEnW6PVKk+FeAJyaa9R4Hfv37zfc50iOj1YAVgCm04rOK6O0ObSkh7PfLkKIj5MCmi8AZjbxecZXOxSZ5cyivkit3qAArABsxhxwYIr86eRRd5J/ADxlK5/ymyEA27fRLmfC//77L5jof/48cW1MGJXF6KxIEzMeXsWKlTBlirFc61DbLS6uDvbs2W2oG3FxcRg1ynj9IEOd8PNmvkpWrFjBcFwEm3322WexfftO99IpHs9+nbvqdQW2A/hVAC5PadKkiRg4UGO2NCw6rnPDOq1WoACcvFdgszzPtGK/fv3x5pst3adsUSHEHm/z2CcA20FMABPImjBkrEiRQrh69aq3Nrx+nzVrVmzatMXdfe71vlBeoACcfAG8e/du1K3rkikU8FRMmzYt4uP3uUcm7hFC+OTd9QfAzE5K4OJhjwcPHoQJE8YH3HnnGyMt1VABOHkCmA6rqlWr4ODB70yZ9507d0GXLu+666onhFjoSwP+AJinyyxXmFAn5dKlSyhW7FUwwMOo8Jdo585dYD2lSBAF4OQJ4AUL5uO99xIBLqAp+9BDD2HPnnhw7jsJGSSzCyH+9UWpzwC2b6NJOUvq2QTp378fPv/ceP0XKmzYsBGGDh3mS79Dfo0CcPID8JUrV1CqVAn8/jvrHxiXdu3ao0eP990VtRdCjPVVu78AJkEPI7afcjTAwZA7+u+/E2U6+doHl+uWLl2GwoWLBHRvMG9SAE5+AO7RozvmzJltyjR74IEHsHv31+5ZRzzWySGEuOFrI34B2L4Kk1Ta5ReCZUhZjtQMyZ49O778coN7MSczVJuqQwE4eQH4wIEDqF69milBG7Rchw4d0a1bd/c52U4I4bLD9TZpAwFwKlue8HEAzzmUc2vBd2EzPNLU2alTZ7z77nve+h7S73mEtnZtoqJTfvWpZMlSGDaMJKCRKwwlPHDgG0MDqFy5Cnr3ZnHM8BT6eGJiKuHkSZbQNi6stEB/T5o0CXwZVMp335xCCL/oX/0GsH0V5qGVy4vvmDGjMWzYUOOjs3m3U6ZMidWr1+KFFxJCsE3Rq5QoCwRigUGDBmLixAmB3Kp7D4uVsWiZm7QSQvidJRQogOmRZt3R7I5O8FeqdOlShuvBOPTlyJET69Z9ofHjKlEWCJUF9u7dizp1apkSccUxZMmSBZs3b3WPeThlK2mUy1fPs7MtAgKwfRVmBrNL+TVuKd96q7VpttYprWiabqVIWcCbBchzxRpHZtDEOtqaOHES+MrgJg2FEAGx4RkBMO9lqJeLy9iM2Fjnwc2cOQtlypT1Zmv1vbKA6Rbo2LEDli1bapreEiVKYO5cl9qB1B3PCEchRECVwAMGsH0VZsL/DudiaEePHkVsbCXTthx84WeFtieeeMI0QypFygLeLEB6WNLEmiX062zcuBkMG3YSgrakECJgojlDALaDmNtoF0Igs1g7HAN9+eWXsXjxUs25pURZwGoLfP/9cVSuHGtKhKGjrx07voOuXV3KGvGruUKIhkbGYwaAGdTBEMsEnzjfHcqXLwsWCDdLmjRpisGDh5ilTulRFtC1AI9Cq1WrglOn6FcyR5555lls2rTZvVAZuZ7puDKUk2sYwPZVmIe2Lmzt27ZtRaNGhn5cElmPOaPMHVWiLGCFBf777z80a9YUW7e61A8z3JQHP05XIcRIo8rNAjD5YVl3paBzhzp1egdLliSqBhFwn7mFXrBgYUSEWgY8SHVjyCwwYEB/TJ48ydT2PVQiIR/VK0KI/4w2ZgqA7aswCw/vBZDAick4aXIG/fZbkqwgfo2BBZ9WrFjl7gzwS4e6WFnA3QILFy7Au+92MdUwmTJl0rbOadI85KyXWUaFhRDfmtGYaQC2g/hDAC4FlMiXS/YCM4UUJAQxPdRKlAWMWoBb5ubNm4FUUWbKggWLUKxYMXeVQ4UQiVKQAm3XbAAzW+kgMyqcO9Sz5/uYNWtmoH3Uva9gwYKYN28B7r//flP1KmXJywLffvst4uJq4+ZN3cIHARvDg9OVwdQvCiFMa8xUANtXYZ4NbwOQUDeFxomNjTEtGNxhVf66zZgxy927F7DR1Y3JywL0NNeqVcO0/F6H9VgHe/XqNe7zku+7pY2c+eo9HdMBbAdxfwAu6SUM8KhatTJu33YpMG54xpCKc+LEyeqM2LAlk5cChkdy5f3ll19MHTjzfNesWQemxbrJAFu0VV9TG3OOoDJTsZSSjixGaCWQ4FH/9OnT0Lv3B2Y2pemqWrUqxowZF9HF0kw3ilLo0QJWgZcNejjq5AkNS6SY+5JtFYDtqzBjxuhpc3HBvfNORyxdusT06VWpUgzGj5+gVmLTLRtdCq0Erwd+a9K2FhBCmBcZ4vRILNlCO/RLKel+nuE8BW7cuKGVZuGW2mwpVaq0VltVObbMtmx06Pvhhx9Qr16c6dtmWoe56yzWxy20mzQTQrhgwExrWgpg+0rMJOU3nTt95sxpLdbUjDKl7sZgvSUSxZPxT4mygMMCpIFt3LiR6Q4r6udxJgkonn76aXeDTxVCuMx9s59IMADMjHxmWxRy7jzP3po2bWJa1pKz7jx58mD69Jkqg8ns2RKh+nbs2I5WrVqCMfpmC2v5MjqwSJFX3FUz2orvvaYdGen13XIA21fhZ2y+pn02MrxHnTsxc+YM9OrV02ybavqYfkgQE8xKkq8FmM/bpUtn04M0HBYdPnwEGC7pJr/ao63Ic2WpBAXAdhCz9No65/Nh/jvJ4VhnyQp58MEHMW7cBJQtqwgBrLBvOOu8c+cOhg8fhrFjx1jWTQ8FuXneGyuE2GBZw06KgwZgO4iZIe1Cw0hDt2nztmGGR0/GSpEiBd5/vyfeeuttCBHU4Qbj+ak2dCzArXKHDu2xYcN6y+xTrVo1jB07Xm9OdRNCuGTmWdYJK4+RPHVaSkneW1Z4SBB6phs0qGdK0XBP7cbGVsaoUR+7U3laaVulOwQW+PHHH9G6dUv8739MUbdGSI3DCEAdgonxQgjypgdNgr4kSSkZYkmiodedR/nXX3+hVq2aYN1Vq4R0JjxmypnzeauaUHpDaIGVK1doNDhWOKscw8qbNy8WLVrsnmHEr1fa8gBqmZEi6I8Jgw5gdk5K+aC93nA+584y7ZCxqTyvs0p4Rty7d180atRIbamtMnKQ9ZLSuF+/Ppg925yyJ566zyy45ctXImPGjO6XHAJQXAhhvNaun7YLCYDtIOahGWsOuxyeXbhwQYtRPX3aWgde+fLlMWLER3oPw08TqstDaYEjR46A7JHksbJSCN5Fi5bgySefdG+GvFEsxm0ef5QfAwkZgO0gZsQ3z4gfd+6zleFuzu3wAJ4gZkKEksiyAHN3WQ3k008/seyIyGGRJMB7EUAJGzGdOTVXAngEIQWwHcTcRm8G4LIvIYgbNXrD8pWYfaBHsX//AXj00ccCMKG6JdgWoIOKZ7uHDjH13FpJAryXAZQVQnD7HDIJOYDtIGaU1kYA6Zwtwe10/fp1LX0ndrTHIsu9en2gHcqr46aQzcckG2YJ21GjPsKUKZ9bvuqyI3R6zp+/UG/b/CeACkIIUkiFVMICwHYQkwiAB3cu0eB0bNWrV9dS77TzE2Bt4n79+uHFF18K6YNRjbta4Isv1qFv3z74+eefg2KafPlexKxZs/Vom1gIu5LZifmBDipsAGwHcQlbCBprdrpkIvCIqUmTRpaeEzsbkCtwnTp10K1bDxVPHejMMuk+OqkGDRqAnTsDLl7gd0+KFy+Ozz+fondURC7nykII5rqHhYQVgJ1WYoLYZTvNYA/S1BqtyeuP1Xnk1LZtO7Rq1RoMy1QSPAucP38eI0YM13LHpQyobFBAnWXAz9ix4/SCNLhtriKE+CogxRbdFHYAtoP4ZRubxxfuyQ8Muxw8eJBlsdOebJw+fXotFJPVEhWQLZqJdrWXLl3CuHFjMXv2LFNLm/jSa/5Q0w9yzz0JdG6O25icwPhmZhiFlYQlgO0gzmU7X/sSADOZXIRZTKTmIaCDKQ4gN23aVG97FcyuRF1bXHHHjh2NhQsXBh24TAn88MOhqF+/gZ5dz9jfea0LETTwNMMWwHYQZ7LlEa8mJYn7GJlP3L59O0tIAbzZM02aNGjQoAFatHgTTz+d2dvl6vskLHD48GFMnTpFK+NpNi+zL4ZnoQCG15IIQkdICcVtc3A8Z7502O2asAawHcQsmraAzgP38ZHZg4naVtDz+GJLbrViYmLQvHkLLaFbHT/5YjVoQF2//ktMnToVX3/NEtOhEdK/TpkyFTzr1RGmvtYVQtBxFbYS9gC2g5gvJaPds5j4HTmnu3fvZglRnj9PLUuWLNoWjMRmKiBE33KMcV+0aCEWL15kCS+VP88rLq4uhgz50BOn+HgAHYKdmOBP/x3XRgSAHZ2VUrIK4lB3UgB+T8paFqcym3faX6NyVSaBQI0aNVGuXPlk7/S6cuWKdnJA0O7dG/K4B6ROnVp73yWAdYTJ+D3MqBro77wJ9PqIArB9NS4PYJ576CW/41a6Xbu2pleACNS4nCxlypTVttkVK1ZKNkR79CR/+eWXWLduDXbv3h2Sd1u9Z5YtWzZMmDAJuXLRP5pIGBrZQAjBiMCIkYgDsB3E9EyTXNqFKM+xpeZKbHYtJqNPlJ7OggULoVSpUihdujTy5MkLsoVEg/Cd9ptv9mPHjh3Yvn0bDhw4ENSzW19syCPAnj17eaIcJl9bbSEEPc4RJREJYDuIWUiNhEe6tJ2sivjuu51NLW1q5pPNkCGDVrnu5ZcLasDOly9fxJDSM6jm8OFD2LdvH/bujddW2WvXwtPXQ3JDMrGUKFHS0+ObAqC91eyRZs4dZ10RC2DHIOzk8XRwpXU3EusTczU2s8i4VQ8iVapUGogZg0vv6PPPP4/cuXOF/LyZ77BkSfn++xM4duwoWM3v6NEjYDX7cBf6IQYPHgImqujIX3ZHlbllM4NslIgHsH01ZhmXOe61mBy23LZtK3r06I5z50KSc23okT711FPaMQfPm0kcfveTGY8+mhE8w3z44fTg9jwQ4db3ypU/QJD+8stFnD9/DgyooJ343wTt5csMQoosyZz5GQwePFjzP3gQnl01sqrcSTCtFRUAtoOYs7g3ABJNJ5rR5ElibO20aVODHsFl9QNlYEmGDI8gTZq78dos73HvvSmRIsXdx3vnjsR///2bwBV19eo1/P77b5ZyR1k9Zj39/CFr3fotdOrU2dO7LouLDSGbsRWFxkIx5qgBsNOWmiXRpwHIqWdQeqqZlrZnD9l8lESLBZgGOnTo0KQIC78H0FwIsStaxsxxRB2A7avx/fbVmDzUuvvLdevWYsCAATh37mw0Pc9kN5ZnnnkWvXr1QuXKVTyNnavuSACsz3sj2gwUlQB2Wo3zA5isd9zEa8hmOHnyJC375erVoBMKRttcCup46Jjq1KkLmjVrlpT3nsdDrYQQjGmOSolqADu9G79j++9+tl9ixlUnEjpxxo8fp0VzkbZFSfhagO/3PNNt06at5sTzIDzT6g/gk2h51/U00KgHsNNqTD7QQQCa2bZTuhEUPHYaPfozLQiEq7OS8LGAA7jMy+YZugdhful0AB8IIS6ET++t60myAbATkJmaOJZcvp7MylBAbq3nzJmtttbWzT2fNJNAoWnTZhqhQhLApS4eDbUTQnzjk+IouSjZAdi+rea469OxAYDc1LrC6KK5c+do+ao8H1USPAvwvJv51szw8lKs/ZTdYTlfCBE87p3gmSLJlpIlgJ1WY3qom3LLZdt6PefJUgx4WL16lfaOvH9/2LGqhMlUMqcbhQsXxptvttISQHSobZwbYemOgQBmRPt7blKWTdYAdgJyKntMNYNAXEq9uBuPJTzmzp2rpcf9+Sd5zpQYtQCpimrWrKWttrlz5/amjluhwQCmCCFuebs42r9XAHZ6wlJKJkg0AdDFlrKYZAlDOrl4lrxgwXzs2rUr6qK7rJ74ZC8pWbKUBtqKFSuCseBehMWPPqWTKhrPc70N3tP3CsA6lpFS0ktdFQAJBMhVnaRcvnwZ69atw4oVyxEf/3XYpdJ563+wvmf6ZNGiRRETEwvStz7+uEtJLE/dIAczAzFWCyGCy2IYLMMYaEcB2IvxpJRFAHQGUBNAam+2vnjxosZAsXHjRi1c89at5L3L48r6yiuvIDa2CmJjY32tBskzvGU2bvCPhRDx3myenL9XAPbx6UspH7Fvr1vZSPa8vqhRLfNmCeItW7Zg8+ZNQSnU5uNwLL2MzBelSpXWPlxxSZDvo5C6dRKAWUIIMmQo8WIBBeAApoiUknWcWrIiu14esieVTNPbt28v4uPjtb/Hjx+P+HdnbotJUVOwYEEUKlRYY+fkEZAfwhhWsqvQKRW8+il+dDCcL1UANvB07E6vWNKP2t+ZdUM1PTVx7dpV7ViK9X+YJcWymadOnQwbDin3fjNdL1u27BrZAEGbP38BFChQAExn9FOuA1gFYCGAdZHKhuHnmC25XAHYJLNKKblPZEpMHZaeBOAx3i+pJsmqeeLECY0F4+zZs/bPGZw5cxYXLvxsObgJUpII8EPigMyZM+O5557TWEJy5MgRMHkAgN8BbACwmAXslCfZnImnAGyOHV20SCnJY03nF1fnGNtqU9BT/LU/zbPIF5MtyIV948bdvzdv/qP95bEWS81IeQe3bt3W1N6+fdeBljLl3SOalCnvRYoU92gE9Pfdd5/9k9r+937tXZUxxyYR1NNjzKgX1rgiSXp8JPAs+/M8wuFaBeAgPAUp5aMAStuKmL9mW4X4/sziw4Hx4AShvwE2wbzb72y7D1bv47vsViFE5PHxBDj4UN2mABwCy0spyX3zih3QTK540RYSmCWCCBYYc/yjzYl3EMABO2C/FkLw3VZJEC2gABxEYyfVlJSSnqC8djAT0Nns8dmM0WaEWCjkpi398if7h0kDBCw/h8O9ZlAojBWKNv8PB1hDEnGgI7EAAAAASUVORK5CYII="
                                            alt="image001.png"
                                            crossorigin="use-credentials"
                                            fetchpriority="high"
                                            class="Do8Zj"
                                            style="min-width: auto; min-height: auto; max-width: 100%; height: auto;"
                                        />
                                    </a>
                                </div>
                                <p></p>
                            </td>
                            <td style="padding: 0in;">
                                <p style="margin: 0in 0in 3.75pt; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="letter-spacing: 0.6pt; font-family: 'High Tower Text', serif; font-size: 11pt;"><b>Debbie Scott</b></span>
                                    <span style="font-family: 'High Tower Text', serif; font-size: 11pt;">
                                        <br />
                                        Purchasing Card Manager
                                    </span>
                                </p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif; color: black !important;">
                                        <b>
                                            <a
                                                href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fhighgate.com%2F&amp;data=05%7C02%7CRGioia%40blueoceansp.ai%7C97e340e09ca242351d6508dc5308d3b4%7Ceb8172eda7a8432cb2620327e4eb2097%7C0%7C0%7C638476544592283358%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&amp;sdata=JwgEpVk6gn84Y37IsTA9SMfyZ2DOu1hBMB4%2FRBacsLY%3D&amp;reserved=0"
                                                target="_blank"
                                                rel="noopener noreferrer"
                                                data-auth="Verified"
                                                originalsrc="https://highgate.com/"
                                                shash="OaiyXwoxdmDTZSd02MPQNPh9h+SD95zQCxLsWpd/82Er/CovfPPLzjUtzYugtUiENP4TD63lrDFJ4BrHQevwogmEqpZPkH91yh2tLJjb94beGLdvwL5Mr8wm23DI7YE1cJRpIUHlEmzrPFBka/O3R9HJ6WzK0cPE1oZLb9wg9EA="
                                                title="Dirección URL original: https://highgate.com/. Haga clic o pulse si confía en este vínculo."
                                                id="OWA6b8195c9-c50b-2c4d-a582-504610158eff"
                                                class="x_OWAAutoLink"
                                                data-loopstyle="linkonly"
                                                style="color: black !important; margin-top: 0px; margin-bottom: 0px;"
                                                data-linkindex="18"
                                            >
                                                HIGHGATE.COM
                                            </a>
                                        </b>
                                        <a
                                            href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fhighgate.com%2F&amp;data=05%7C02%7CRGioia%40blueoceansp.ai%7C97e340e09ca242351d6508dc5308d3b4%7Ceb8172eda7a8432cb2620327e4eb2097%7C0%7C0%7C638476544592289193%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&amp;sdata=DyErtVcFKDi%2FPsihKOeJjy2TOqhwREqr2EEFoG6OSLY%3D&amp;reserved=0"
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            data-auth="Verified"
                                            originalsrc="https://highgate.com/"
                                            shash="Nm8OxrUYVK3FGLRRiJschQDTUH5C5O011+G9A5pthzu0xClLov/iVl1CNwEfRq2XyJteOBUGIzAw3EXJpRfcrum53jzg7GeVqZAcJoKeOHykUzwZTF+SVxkIm2w57UX1BDhpEFpb5qNGPK6kbz4QLhPshFBQFxyMebhSKLAbODQ="
                                            title="Dirección URL original: https://highgate.com/. Haga clic o pulse si confía en este vínculo."
                                            id="OWA75d8f146-8bc6-a384-9b72-b1498189cf37"
                                            class="x_OWAAutoLink"
                                            data-loopstyle="linkonly"
                                            style="color: black !important; margin-top: 0px; margin-bottom: 0px;"
                                            data-linkindex="19"
                                        >
                                        </a>
                                    </span>
                                    <span style="font-family: 'High Tower Text', serif;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
                                </p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif; color: rgb(95, 95, 95) !important;">O:&nbsp; </span>
                                    <span style="font-family: 'High Tower Text', serif; color: black !important;">
                                        <b>
                                            <a
                                                href="tel:972-842-9818"
                                                target="_blank"
                                                rel="noopener noreferrer"
                                                data-auth="NotApplicable"
                                                title="Dial 972-842-9818"
                                                id="OWA069696d9-1fa8-293f-ad94-ac690be866c2"
                                                class="x_OWAAutoLink"
                                                data-loopstyle="linkonly"
                                                style="color: black !important; margin-top: 0px; margin-bottom: 0px;"
                                                data-linkindex="20"
                                            >
                                                (972) 777-4385
                                            </a>
                                        </b>
                                    </span>
                                </p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif;">C:&nbsp; <b>(940) 255-9411</b>&nbsp;&nbsp;&nbsp;</span>
                                </p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;"><span style="font-family: 'High Tower Text', serif; font-size: 9pt;">&nbsp;</span></p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif; font-size: 9pt;">
                                        email:
                                        <a href="mailto:Pcard@higHgate.com" id="OWA6654b3ab-7bae-d671-ac16-4a24b5a2db0e" class="x_OWAAutoLink" data-loopstyle="linkonly" style="margin-top: 0px; margin-bottom: 0px;" data-linkindex="21">Pcard@higHgate.com</a>
                                    </span>
                                </p>
                            </td>
                        </tr>
                    </tbody>
                </table>
            """
        recipientsTo = dataset['recipients']
        recipientsCC = dataset['recipientsCC']
        attachments = []
        # Catch failures to prevent process abortion if only some reports cannot be sent out.
        try:
            # Prepare Attachments
            data = open(dataset['fpaths'], "rb").read()
            file_encoded = base64.b64encode(data).decode('UTF-8')
            attachments.append({
                    "ContentBytes": file_encoded,
                    "Name": dataset['fpaths'].name
                })
            if not attachments:
                raise Exception(f"No files to send for: {dataset}")
        
            status = f"Email Sent Successfully to {recipientsTo}."
        except:
            logger.error(f"{traceback.format_exc(chain=True)}")
            failed_datasets.append(dataset)
            status = f"Failed to Send email for {recipientsTo}."
        logger.info(status)

    # Send email to Highgate to inform about unsent reports.
    if failed_datasets:
        import pandas as pd
        df = pd.DataFrame(data=failed_datasets).fillna(' ').T
        failed_datasets_formatted = df.to_html()
        failed_email_subject = f"STR Reports - Unsent Files"
        failed_email_body = f"The following reports could not be sent out:\n\n {failed_datasets_formatted} \n\nPlease check the reports on SharePoint for manual distribution."

        
def send_auxiliary_email(emailing_dataset: list, auxiliary_email: str, start_date: str, end_date: str):
    failed_datasets = []
    start_date_obj = datetime.strptime(start_date,"%m.%d.%y")
    next_day_date_obj = start_date_obj + relativedelta(months=1)
    next_day_date_formatted = next_day_date_obj.strftime("%m.%d.%y")

    for dataset in emailing_dataset:
        # Prepare Email
        email_subject = f"{start_date} to {end_date} - BI-WEEKLY Pcard Outstanding"
        email_body = f"""
                <div style="font-family:Calibri, Helvetica, sans-serif;">
                    <p>Dear VP of Finance and Property GMs,</p>
                    <p>Please see the attached Bi-Weekly Pcard Compliance report for {start_date} to {end_date}.  This report summarizes the Pcard bi-weekly spend for your properties and shows transactions that still need to be coded, have receipts uploaded, and the sign-off completed (highlighted on the report). Per policy, these actions are due as soon as the transactions are posted in Works.</p>
                    <p>The first tab contains pivot tables summarizing the spending for your properties and the 2nd tab contains all of the transactional information. We hope you will contact your properties and ensure they are coding the transactions, uploading the receipts, and signing off so that we have all transactions coded by month's end. Pcard limits may be taken to $1.00 on {next_day_date_formatted} for those who have not coded their transactions for {calendar.month_name[int(start_date[:2])]}.</p>
                    <p>This report was pulled today {datetime.now().strftime("%m.%d.%y")}, if the coding has been completed, the receipts have been uploaded and sign-offs are done, please disregard.</p>
                    <p>Please let me know if you have any questions. Thanks.</p>
                    <p>EMAIL: <a href="mailto:PCARD@HIGHGATE.COM">PCARD@HIGHGATE.COM</a><br></p>
                </div>
                <table style="width: 471pt; border-collapse: collapse; border-spacing: 0px; box-sizing: border-box;">
                    <tbody>
                        <tr>
                            <td style="padding: 0in 11.25pt 0in 0in; width: 67.75pt;">
                                <p align="center" style="line-height: 12pt; margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif; color: blue !important;">
                                        <a
                                            href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.highgate.com%2F&amp;data=05%7C02%7CRGioia%40blueoceansp.ai%7C97e340e09ca242351d6508dc5308d3b4%7Ceb8172eda7a8432cb2620327e4eb2097%7C0%7C0%7C638476544592277339%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&amp;sdata=8wiKkcob3NbPbz%2BCnMZ7jQywv8ssNI4dNRSDHnu4BQM%3D&amp;reserved=0"
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            data-auth="Verified"
                                            originalsrc="https://www.highgate.com/"
                                            shash="QI62XVEPKAdFZpNbtFY4UjbHykAS46sHsk3uIzhD3PmmGWd8AYzUSy7yonmr/tmAkIEVMjn6NNR+/19pbWTn9PMIUiO8geHXwksY6UhlKZ5dZoEk9s7gb9zH+P8JGq5KpA3TaksfKwb0aGFI7iCkUMuXe3D/D+tudDKnslV3zZQ="
                                            title="Dirección URL original: https://www.highgate.com/. Haga clic o pulse si confía en este vínculo."
                                            id="OWAf9708882-463d-4f12-5b34-a9fb65f6ad50"
                                            class="x_OWAAutoLink"
                                            data-loopstyle="linkonly"
                                            style="color: blue !important; margin-top: 0px; margin-bottom: 0px;"
                                            data-linkindex="16"
                                        ></a>
                                    </span>
                                </p>
                                <div>
                                    <a
                                        href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.highgate.com%2F&amp;data=05%7C02%7CRGioia%40blueoceansp.ai%7C97e340e09ca242351d6508dc5308d3b4%7Ceb8172eda7a8432cb2620327e4eb2097%7C0%7C0%7C638476544592277339%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&amp;sdata=8wiKkcob3NbPbz%2BCnMZ7jQywv8ssNI4dNRSDHnu4BQM%3D&amp;reserved=0"
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        data-auth="Verified"
                                        originalsrc="https://www.highgate.com/"
                                        shash="QI62XVEPKAdFZpNbtFY4UjbHykAS46sHsk3uIzhD3PmmGWd8AYzUSy7yonmr/tmAkIEVMjn6NNR+/19pbWTn9PMIUiO8geHXwksY6UhlKZ5dZoEk9s7gb9zH+P8JGq5KpA3TaksfKwb0aGFI7iCkUMuXe3D/D+tudDKnslV3zZQ="
                                        title="Dirección URL original: https://www.highgate.com/. Haga clic o pulse si confía en este vínculo."
                                        id="OWAf9708882-463d-4f12-5b34-a9fb65f6ad50"
                                        class="x_OWAAutoLink"
                                        data-loopstyle="linkonly"
                                        style="color: blue !important; margin-top: 0px; margin-bottom: 0px;"
                                        data-linkindex="17"
                                    >
                                        <img
                                            data-imagetype="AttachmentByCid"
                                            data-custom="AAMkAGE0NDhmNmY4LWI5OTQtNGYxNy05MmFiLTE1YmY1MzNlNDBhNQBGAAAAAADMrIwFplGzS7yzyq7lAXxaBwAZmXyA7SUvQ72Ye0QZ%2FmlMAAAAAAEMAAAZmXyA7SUvQ72Ye0QZ%2FmlMAACKdmpdAAABEgAQAKiWDGowO39Om6ghceLYjrg%3D"
                                            naturalheight="0"
                                            naturalwidth="0"
                                            src="data:image/png;base64,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"
                                            alt="image001.png"
                                            crossorigin="use-credentials"
                                            fetchpriority="high"
                                            class="Do8Zj"
                                            style="min-width: auto; min-height: auto; max-width: 100%; height: auto;"
                                        />
                                    </a>
                                </div>
                                <p></p>
                            </td>
                            <td style="padding: 0in;">
                                <p style="margin: 0in 0in 3.75pt; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="letter-spacing: 0.6pt; font-family: 'High Tower Text', serif; font-size: 11pt;"><b>Debbie Scott</b></span>
                                    <span style="font-family: 'High Tower Text', serif; font-size: 11pt;">
                                        <br />
                                        Purchasing Card Manager
                                    </span>
                                </p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif; color: black !important;">
                                        <b>
                                            <a
                                                href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fhighgate.com%2F&amp;data=05%7C02%7CRGioia%40blueoceansp.ai%7C97e340e09ca242351d6508dc5308d3b4%7Ceb8172eda7a8432cb2620327e4eb2097%7C0%7C0%7C638476544592283358%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&amp;sdata=JwgEpVk6gn84Y37IsTA9SMfyZ2DOu1hBMB4%2FRBacsLY%3D&amp;reserved=0"
                                                target="_blank"
                                                rel="noopener noreferrer"
                                                data-auth="Verified"
                                                originalsrc="https://highgate.com/"
                                                shash="OaiyXwoxdmDTZSd02MPQNPh9h+SD95zQCxLsWpd/82Er/CovfPPLzjUtzYugtUiENP4TD63lrDFJ4BrHQevwogmEqpZPkH91yh2tLJjb94beGLdvwL5Mr8wm23DI7YE1cJRpIUHlEmzrPFBka/O3R9HJ6WzK0cPE1oZLb9wg9EA="
                                                title="Dirección URL original: https://highgate.com/. Haga clic o pulse si confía en este vínculo."
                                                id="OWA6b8195c9-c50b-2c4d-a582-504610158eff"
                                                class="x_OWAAutoLink"
                                                data-loopstyle="linkonly"
                                                style="color: black !important; margin-top: 0px; margin-bottom: 0px;"
                                                data-linkindex="18"
                                            >
                                                HIGHGATE.COM
                                            </a>
                                        </b>
                                        <a
                                            href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fhighgate.com%2F&amp;data=05%7C02%7CRGioia%40blueoceansp.ai%7C97e340e09ca242351d6508dc5308d3b4%7Ceb8172eda7a8432cb2620327e4eb2097%7C0%7C0%7C638476544592289193%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&amp;sdata=DyErtVcFKDi%2FPsihKOeJjy2TOqhwREqr2EEFoG6OSLY%3D&amp;reserved=0"
                                            target="_blank"
                                            rel="noopener noreferrer"
                                            data-auth="Verified"
                                            originalsrc="https://highgate.com/"
                                            shash="Nm8OxrUYVK3FGLRRiJschQDTUH5C5O011+G9A5pthzu0xClLov/iVl1CNwEfRq2XyJteOBUGIzAw3EXJpRfcrum53jzg7GeVqZAcJoKeOHykUzwZTF+SVxkIm2w57UX1BDhpEFpb5qNGPK6kbz4QLhPshFBQFxyMebhSKLAbODQ="
                                            title="Dirección URL original: https://highgate.com/. Haga clic o pulse si confía en este vínculo."
                                            id="OWA75d8f146-8bc6-a384-9b72-b1498189cf37"
                                            class="x_OWAAutoLink"
                                            data-loopstyle="linkonly"
                                            style="color: black !important; margin-top: 0px; margin-bottom: 0px;"
                                            data-linkindex="19"
                                        >
                                        </a>
                                    </span>
                                    <span style="font-family: 'High Tower Text', serif;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>
                                </p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif; color: rgb(95, 95, 95) !important;">O:&nbsp; </span>
                                    <span style="font-family: 'High Tower Text', serif; color: black !important;">
                                        <b>
                                            <a
                                                href="tel:972-842-9818"
                                                target="_blank"
                                                rel="noopener noreferrer"
                                                data-auth="NotApplicable"
                                                title="Dial 972-842-9818"
                                                id="OWA069696d9-1fa8-293f-ad94-ac690be866c2"
                                                class="x_OWAAutoLink"
                                                data-loopstyle="linkonly"
                                                style="color: black !important; margin-top: 0px; margin-bottom: 0px;"
                                                data-linkindex="20"
                                            >
                                                (972) 777-4385
                                            </a>
                                        </b>
                                    </span>
                                </p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif;">C:&nbsp; <b>(940) 255-9411</b>&nbsp;&nbsp;&nbsp;</span>
                                </p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;"><span style="font-family: 'High Tower Text', serif; font-size: 9pt;">&nbsp;</span></p>
                                <p style="margin: 0in; font-family: Calibri, sans-serif; font-size: 10pt;">
                                    <span style="font-family: 'High Tower Text', serif; font-size: 9pt;">
                                        email:
                                        <a href="mailto:Pcard@higHgate.com" id="OWA6654b3ab-7bae-d671-ac16-4a24b5a2db0e" class="x_OWAAutoLink" data-loopstyle="linkonly" style="margin-top: 0px; margin-bottom: 0px;" data-linkindex="21">Pcard@higHgate.com</a>
                                    </span>
                                </p>
                            </td>
                        </tr>
                    </tbody>
                </table>
            """
        recipientsTo = [auxiliary_email]
        attachments = []
        # Catch failures to prevent process abortion if only some reports cannot be sent out.
        try:
            # Prepare Attachments
            data = open(dataset['fpaths'], "rb").read()
            file_encoded = base64.b64encode(data).decode('UTF-8')
            attachments.append({
                    "ContentBytes": file_encoded,
                    "Name": dataset['fpaths'].name
                })
            if not attachments:
                raise Exception(f"No files to send for: {dataset}")

            status = f"Email Sent Successfully to {recipientsTo}."
        except:
            logger.error(f"{traceback.format_exc(chain=True)}")
            failed_datasets.append(dataset)
            status = f"Failed to Send email for {recipientsTo}."
        logger.info(status)

    # Send email to Highgate to inform about unsent reports.
    if failed_datasets:
        import pandas as pd
        df = pd.DataFrame(data=failed_datasets).fillna(' ').T
        failed_datasets_formatted = df.to_html()
        failed_email_subject = f"STR Reports - Unsent Files"
        failed_email_body = f"The following reports could not be sent out:\n\n {failed_datasets_formatted} \n\nPlease check the reports on SharePoint for manual distribution."


def send_email_with_multiple_attachments(auxiliary_email: str, file_attachments: list):
    failed_datasets = []

    # Prepare Email
    email_subject = f"EOM Pcard monthly report"
    email_body = f"""<p>This is an automatically processed email.</p>
    <p>Files processed based on the: <strong>{email_subject}</strong> generated for today: <strong>{datetime.now().strftime("%m.%d.%y")}</strong>.</p>"""

    if auxiliary_email:
        logger.info(f"Auxiliary email was specified. Now triggering email to the following email box: {auxiliary_email}")
        recipientsTo = [auxiliary_email]
    else:
        logger.info(f"Now triggering email to the configured email box: {config.Emails.eom_pcard_default_email}")
        recipientsTo = [config.Emails.eom_pcard_default_email]
    attachments = []
    # Catch failures to prevent process abortion if only some reports cannot be sent out.
    try:
        # Prepare Attachments
        for file_attachment in file_attachments:
            data = open(file_attachment, "rb").read()
            file_encoded = base64.b64encode(data).decode('UTF-8')
            attachments.append({
                    "ContentBytes": file_encoded,
                    "Name": file_attachment.name
                })
        if not attachments:
            raise Exception(f"No files to send for EOMPcard Report.")
        send_email(
            emailSubject=email_subject,
            emailBody=email_body ,
            recipientsTo= recipientsTo,
            attachments=attachments,
            special_sender="highgate"
        )
        status = f"Email Sent Successfully to {recipientsTo}."
    except:
        logger.error(f"{traceback.format_exc(chain=True)}")
        failed_datasets.append("EOMPcard report")
        status = f"Failed to Send email for {recipientsTo}."
    logger.info(status)

    # Send email to Highgate to inform about unsent reports.
    if failed_datasets:
        import pandas as pd
        df = pd.DataFrame(data=failed_datasets).fillna(' ').T
        failed_datasets_formatted = df.to_html()
        failed_email_subject = f"STR Reports - Unsent Files"
        failed_email_body = f"The following reports could not be sent out:\n\n {failed_datasets_formatted} \n\nPlease check the reports on SharePoint for manual distribution."
        send_email(
            emailSubject=failed_email_subject,
            emailBody=failed_email_body,
            recipientsTo=Emails.recipients_alternate,
            special_sender="highgate"
        )