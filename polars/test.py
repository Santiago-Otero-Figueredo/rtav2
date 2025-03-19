from app.auxiliares import cruce_sistecredito, cruce_oms_mercado_pago_clase, cruce_addi_erp
from app.lectores.modelos import ConfiguracionLector
from app.lectores.lector_oms import LectorOMS
from app.lectores.lector_mercadopago import LectorMercadoPago
from app.lectores.lector_erp import LectorERP
from app.lectores.lector_addi import LectorADDI
from app.lectores.lector_mercadolibre import LectorMercadoLibre
from app.lectores.sistecredito.lector_facturas import LectorSisCredFacturas
from app.lectores.sistecredito.lector_pagare import LectorSisCredPagare

#########################################################################################################
############################# USAR LA RUTA DE insumos_test PARA LAS PRUEBAS #############################
#########################################################################################################

from app.settings import (RUTA_RAIZ,
                            RUTA_INSUMOS_ADDI,
                            RUTA_INSUMOS_ERP,
                            RUTA_INSUMOS_MP,
                            RUTA_INSUMOS_ML,
                            RUTA_INSUMOS_OMS,
                            RUTA_INSUMOS_FACTURA_SISTECREDITO,
                            RUTA_INSUMOS_PAGARE_SISTECREDITO,
                            RUTA_RESULTADOS)





def prueba():

    #prueba_oms()
    #prueba_addi()
    prueba_sistecredito()


def prueba_oms():

    config = ConfiguracionLector(ruta_archivo=f'{RUTA_INSUMOS_ERP}/ERP.xls')
    lector_erp = LectorERP(config)
    df_erp = lector_erp.dataframe()

    config = ConfiguracionLector(ruta_archivo=f'{RUTA_INSUMOS_MP}/MERCADOPAGO 172.xlsx')
    lector_mercadopago = LectorMercadoPago(config)
    df_mp = lector_mercadopago.dataframe()

    config = ConfiguracionLector(ruta_carpeta=f'{RUTA_INSUMOS_ML}/')
    lector_mercadolibre = LectorMercadoLibre(config)
    df_mercadolibre = lector_mercadolibre.dataframe()

    config = ConfiguracionLector(ruta_carpeta=f'{RUTA_INSUMOS_OMS}/')
    lector_oms = LectorOMS(config)
    df_oms = lector_oms.dataframe()


    cruce_oms_mercado_pago_clase(df_oms, df_mercadolibre, df_mp, df_erp, RUTA_RESULTADOS)

def prueba_addi():
    config = ConfiguracionLector(ruta_archivo=f'{RUTA_INSUMOS_ERP}/ZOMAC de ADDD.xls')
    lector_erp = LectorERP(config)
    df_erp = lector_erp.dataframe()

    config = ConfiguracionLector(ruta_carpeta=f'{RUTA_INSUMOS_ADDI}/')
    lector_addi = LectorADDI(config)
    df_addi = lector_addi.dataframe()

    cruce_addi_erp(df_erp, df_addi, RUTA_RESULTADOS)


def prueba_sistecredito():
    config = ConfiguracionLector(ruta_archivo=f'{RUTA_INSUMOS_ERP}/ERP_3_12_25.xls')
    lector_erp = LectorERP(config)
    df_erp = lector_erp.dataframe()

    config = ConfiguracionLector(ruta_carpeta=f'{RUTA_INSUMOS_OMS}/')
    lector_oms = LectorOMS(config)
    df_oms = lector_oms.dataframe()

    config = ConfiguracionLector(ruta_archivo=f'{RUTA_INSUMOS_FACTURA_SISTECREDITO}/Facturas_pagadas_3_12_5.xlsx')
    lector_facturas_sistecredito = LectorSisCredFacturas(config)
    df_factura = lector_facturas_sistecredito.dataframe()

    config = ConfiguracionLector(ruta_archivo=f'{RUTA_INSUMOS_PAGARE_SISTECREDITO}/PAGARES SISTECREDITO.xlsx')
    lector_pagares_sistecredito = LectorSisCredPagare(config)
    df_pagare = lector_pagares_sistecredito.dataframe()

    cruce_sistecredito(df_oms, df_factura, df_pagare, df_erp, RUTA_RESULTADOS)



prueba()
