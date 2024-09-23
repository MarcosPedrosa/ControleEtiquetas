package api.poc

import grails.converters.JSON
import groovy.sql.Sql
import org.apache.commons.lang3.time.DateUtils
import java.text.ParseException

class PocController {

    def dataSource

    def index() {
        def sql = new Sql(dataSource)

        if (!params.cFields) {
            render("cFields is required")
        } else {
            try {
                DateUtils.parseDate(params.cFields, "yyyy-mm-dd")
                render([result: sql.rows("SELECT NOME FROM GFERIADO WHERE CODCALENDARIO = '01' AND DIAFERIADO = ?", [params.cFields])] as JSON)
            } catch (ParseException e) {
                e.printStackTrace()
                render("cFields is not valid date. Try use the mask yyyy-mm-dd")
            } catch (Exception e) {
                e.printStackTrace()
                render("Unexpected error")
            }
        }
    }

    def test() {
        render("Resposta OK")
    }

    def test2() {
        def sql = new Sql(dataSource)
        render([result: sql.rows("SELECT TOP 10 NOME FROM GFERIADO WHERE CODCALENDARIO = '01'")] as JSON)
    }
}
