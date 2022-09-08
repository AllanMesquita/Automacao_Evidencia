import azure.functions as func
import psycopg2
import traceback


def main(req: func.HttpRequest) -> func.HttpResponse:
    # logging.info('Python HTTP trigger function processed a request.')

    try:
        import datetime

        # Update connection string information
        host = "psql-itlatam-logisticcontrol.postgres.database.azure.com"
        dbname = "logistic-control"
        user = "logisticpsqladmin@psql-itlatam-logisticcontrol"
        password = "EsjHSrS69295NzHu342ap6P!N"
        sslmode = "require"
        # Construct connection string
        conn_string = "host={0} user={1} dbname={2} password={3} sslmode={4}".format(host, user, dbname, password,
                                                                                     sslmode)
        conn = psycopg2.connect(conn_string)
        print("Connection established")
        cursor = conn.cursor()

        req_body = req.get_json()

        cursor.close()
        conn.close()

        if req_body:
            return func.HttpResponse(
                f"{req_body, 'teste'}",
                status_code=200,
            )
        else:
            return func.HttpResponse(
                "Error",
                status_code=500
            )

    # return "teste"

    except:
        return func.HttpResponse(traceback.format_exc())

    # return func.HttpResponse("retorno", status_code=200)
