require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'
require 'helpers/template_test_helper'

class TestProblems < Test::Unit::TestCase
  include TemplateTestHelper

  def test_active_directory_template
    path_to_template = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'active_directory_woes_for_if_test.docx')
    path_to_correct_render = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'correct_render', 'active_directory_woes_for_if_test.docx')

    render_params = {'fields' => {"Company__Contact___Contract_Info"=>
      [{"Company_Info"=>"",
        "Contract_Info"=>"",
        "Contact_Info"=>[{"Customer_Contact_Info"=>"", "Outgoing_Provider_Contact_Info"=>"", "TLG_Contact_Info"=>""}]}],
     "Active_Directory_Information"=>
      [{"Active_Directory_Domains"=>
         [{"Active_Directory_Domain_Name__FQDN_"=>"Pol",
           "Active_Directory_Domain_Name__NetBIOS_"=>"The",
           "Active_Directory_Domain_Administrator_Username"=>"Gigi",
           "Active_Directory_Domain_Administrator_Password"=>"ghh",
           "Directory_Services_Restore_Mode_Password"=>"ghh",
           "Service_Accounts"=>[{"Service_Account_Username"=>"Uhh", "Service_Account_Password"=>"yhh", "Service_Account_Notes_Description"=>"Guy"}]},
          {"Active_Directory_Domain_Name__FQDN_"=>"Yet"}]}],
     "Internet_Service_Providers"=>"",
     "Internet_Domains"=>"",
     "Web_Hosting"=>"",
     "E_Mail_Server_Provider_Information"=>"",
     "SSL_Certificates"=>"",
     "Remote_Access"=>"",
     "Backup___Disaster_Recovery"=>"",
     "Endpoint_Security"=>"",
     "Wireless"=>"",
     "Software___Licensing"=>"",
     "Hardware___Equipment"=>""}
    }
    check_template(path_to_template, path_to_correct_render, {render_params: render_params})
  end

  def test_solar_pv_template
    path_to_template = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'solar_pv.docx')
    path_to_correct_render = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'correct_render', 'solar_pv.docx')

    render_params = {'fields' => {"Building_Name"=>"First Option", "Building_Location"=>"Somewhere", "Building_Address"=>"99 lolpie lane", "_1__Client_Details"=>[{"_1_1_1__Name___Facilities"=>"Aa", "_1_1_2__Position___Facilities"=>"Aa", "_1_1_3__Telephone___Facilities"=>"74", "_1_2_1__Name___Maintenance"=>"Aha", "_1_2_2__Position___Maintenance"=>"Aha", "_1_2_3__Telephone___Maintenance"=>"5454", "_1_3_1__Name___Building_Owner_Manager"=>"Bags", "_1_3_2__Position___Building_Owner_Manager"=>"Baja", "_1_3_3__Telephone___Building_Owner_Manager"=>"8454"}], "_2__Is_the_building_owned_by_City_of_Cape_Town_"=>"Yes", "_2_1__Comment"=>"Has", "_3__Parapit_Wall"=>[{"_3_1__Parapit_Wall_Height"=>"200 - 600 mm", "_3_2__Comments___Concrete_Roof"=>"Is", "_3_3__Please_take_a_photo_of_the_parapit_wall"=>"IMAGE"}], "_4__Roof_Segment"=>[{"Roof_Segment"=>"First Option", "Roof_Segment_Location"=>"UOu KNow", "_4_1_1__Access_to_roof"=>"Ladder", "_4_1_2__Description_of_access"=>"Is", "_4_1_3__Please_take_a_photo_of_the_access"=>"IMAGE", "_4_2_1_1__Dimensions"=>"Kiss", "_4_2_1_2__Sketch_roof_outline_and_dimension"=>"IMAGE", "_4_2_2__Orientation"=>"25", "_4_2_3__Roof_Tilt_Angle"=>"20 deg", "_4_2_4__Please_take_a_photo_of_the_roof"=>"IMAGE", "_4_3__Comments___Dimensions"=>"Ha", "_4_4_1__Lightning_Protection_Type"=>"Air Terminals", "_4_4_2__Please_take_a_photo_of_the_lightning_protection"=>"IMAGE", "_4_4_3__Air_Terminal_Height"=>"64", "_4_4_4__Please_take_a_photo_of_the_air_terminal"=>"IMAGE", "_4_4_5__Comments___Lightning_Protection"=>"Has", "_4_5_1__Shading_7_10am"=>"Severe", "_4_5_2__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_3__Shading_10am___2pm"=>"Moderate", "_4_5_4__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_5__Shading_2_6pm"=>"Moderate", "_4_5_6__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_7__Comments___Shading"=>"Uses"}, {"Roof_Segment"=>"First Option", "Roof_Segment_Location"=>"UOu KNow", "_4_1_1__Access_to_roof"=>"Scaffolding", "_4_1_2__Description_of_access"=>"It's", "_4_1_3__Please_take_a_photo_of_the_access"=>"IMAGE", "_4_2_1_1__Dimensions"=>"Aha", "_4_2_1_2__Sketch_roof_outline_and_dimension"=>"IMAGE", "_4_2_2__Orientation"=>"87", "_4_2_3__Roof_Tilt_Angle"=>"25 - 40 deg", "_4_2_4__Please_take_a_photo_of_the_roof"=>"IMAGE", "_4_3__Comments___Dimensions"=>"Aha", "_4_4_1__Lightning_Protection_Type"=>"Other -", "_4_4_2__Please_take_a_photo_of_the_lightning_protection"=>"IMAGE", "_4_4_5__Comments___Lightning_Protection"=>"Aha", "_4_5_1__Shading_7_10am"=>"Moderate", "_4_5_2__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_3__Shading_10am___2pm"=>"Moderate", "_4_5_4__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_5__Shading_2_6pm"=>"Moderate", "_4_5_6__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_7__Comments___Shading"=>"Shaun "}], "_5__Inverter"=>[{"_5_1__Inverter_Location"=>"Ha", "_5_2__Wall_space_for_inverters___Width"=>"54", "_5_3__Wall_space_for_inverters___Height"=>"644", "_5_4__Please_take_a_photo_of_inverter_location"=>"IMAGE", "_5_5__Comments___Inverters"=>"Hash"}, {"_5_1__Inverter_Location"=>"Ll", "_5_2__Wall_space_for_inverters___Width"=>"67", "_5_3__Wall_space_for_inverters___Height"=>"76", "_5_4__Please_take_a_photo_of_inverter_location"=>"IMAGE", "_5_5__Comments___Inverters"=>"Zhjs"}], "_6__Tie_in_DB"=>[{"_6_1_1__Tie_in_DB_Location"=>"Ha", "_6_1_2__DB_Name"=>"As", "_6_1_3__Main_or_Sub_DB"=>"Sub DB", "_6_1_4__Please_take_a_photo_of_DB_Location"=>"IMAGE", "_6_2_1__Main_Breaker_Size__A_"=>"77", "_6_2_2__Please_take_a_photo_of_the_main_breaker"=>"IMAGE", "_6_3_1__Fault_Level__kA_"=>"67", "_6_3_2__Please_take_a_photo_of_fault_level"=>"IMAGE", "_6_4_1__Spare_space_in_DB"=>"As", "_6_4_2__Please_take_photo_of_the_spare_DB_space"=>"IMAGE", "_6_5_1__Wall_space_for_new_DB___Width"=>"67", "_6_5_2__Wall_space_for_new_DB___Height"=>"94", "_6_5_3__Please_take_a_photo_of_the_available_wall_space"=>"IMAGE", "_6_6__Comments___Tie_in_DB"=>"Has"}, {"_6_1_1__Tie_in_DB_Location"=>"Aha", "_6_1_2__DB_Name"=>"Kayla", "_6_1_3__Main_or_Sub_DB"=>"Sub DB", "_6_1_4__Please_take_a_photo_of_DB_Location"=>"IMAGE", "_6_2_1__Main_Breaker_Size__A_"=>"976", "_6_2_2__Please_take_a_photo_of_the_main_breaker"=>"IMAGE", "_6_3_1__Fault_Level__kA_"=>"84543", "_6_3_2__Please_take_a_photo_of_fault_level"=>"IMAGE", "_6_4_1__Spare_space_in_DB"=>"Have ", "_6_4_2__Please_take_photo_of_the_spare_DB_space"=>"IMAGE", "_6_5_1__Wall_space_for_new_DB___Width"=>"76", "_6_5_2__Wall_space_for_new_DB___Height"=>"87", "_6_5_3__Please_take_a_photo_of_the_available_wall_space"=>"IMAGE", "_6_6__Comments___Tie_in_DB"=>"Has"}], "_7__Cable_Routing"=>[{"_7_1_1__Is_there_a_clear_path_between_roof_and_DB_room_"=>"true", "_7_1_2__Description_of_path"=>"Ajin", "_7_1_3__Please_take_a_photo_of_path"=>"IMAGE", "_7_2_1__Existing_Wireways"=>"Other -", "_7_2_2__Please_take_a_photo_of_the_existing_wireways"=>"IMAGE", "_7_3_1__Are_additional_wireways_required_"=>"true", "_7_3_2__Please_take_a_photo_where_wireways_are_required"=>"IMAGE", "_7_4__Comments___Cable_Routing"=>"Jana"}], "_8__Contractor_Site_Camp"=>[{"_8_1_1__Location"=>"lat=-33.863409, long=18.641795, alt=165.587158, hAccuracy=65.000000, vAccuracy=10.000000, timestamp=2016-10-04T09:44:57Z", "_8_1_2__Is_there_sufficient_space_for_2x_containers_"=>"true", "_8_1_3__Please_take_a_photo_of_site_camp_location"=>"IMAGE", "_8_2_1__Security"=>"Difficult to secure", "_8_2_2__Please_take_a_photo_of_security"=>"IMAGE", "_8_3__Comments___Contractor_Site_Camp"=>"An"}], "Detail_Sketches"=>[{"Sketch"=>"IMAGE", "Caption_Sketch"=>"He"}, {"Sketch"=>"IMAGE", "Caption_Sketch"=>"Hash"}], "General_Commentary"=>"Shhsxj", "Photo_of_Building_Side_1"=>"IMAGE", "Photo_of_Building_Side_2"=>"IMAGE", "Photo_of_Building_Side_3"=>"IMAGE", "Photo_of_Building_Side_4"=>"IMAGE"}}

      check_template(path_to_template, path_to_correct_render, {render_params: render_params})
  end

  def test_header_and_footer_image
    path_to_template = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'footer_image.docx')
    path_to_correct_render = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'correct_render', 'footer_image.docx')
    render_params = {
      'fields' => {
        'image' => image_to_test_with
      }
    }
    check_template(path_to_template, path_to_correct_render, {render_params: render_params})
  end

  # !!Commented out for now because it is too slow!!
  #
  #
  # def test_plantilla
  #   path_to_template = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'Plantilla.docx')
  #   path_to_correct_render = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'correct_render', 'Plantilla.docx')
  #   render_params = {
  #     'fields' => {"Datos_Servicio"=>[{"Nombre_del_Hospital"=>"IMSS T1 LEON GTO", "Numero_de_Control"=>"1", "Dia_y_Hora"=>"2016-10-28T09:12:31-0500", "Localizaci_n_Hospital"=>"lat=21.14085038, long=-101.68718839, alt=1853.0, accuracy=17.0"}], "Tipo_de_Servicio"=>"Puesta en Marcha", "_rea"=>[{"_rea_a_revisar"=>"Terapia 1", "Existe_Codigo_QR_"=>"true", "Codigo_QR"=>"Cliente:  SOTO REYES ALEJANDRA\nNumero de Control:  1607217935\n\nDATOS DEL TABLERO\nModelo:  HIDP-2  \t# de Serie: 48694\tWO:  220503\nDWG:  17053-A\tPanel:  HIDP2DB012CFNF01-DSU\nInt Gral:   15 Amp\tPolos:    2\tMarca:    C.H.\n# de Ints deriv:   6              \tPolos:    2\tMarca:    C.H.\tAmps:  20  \n\n\nDATOS DEL LIM\n# de Serie:  27517\tModelo:  MARK V\tDate Code:  0916\n\t\t\n\t\t\n\nDATOS DEL TRANSFORMADOR\nCapacidad:   2 KVA    \t# Catalogo:  21-0244L\nPrimario:    220 V   \tSecundario:    120 V\n# de Serie:  P162530135\tImpedancia :   2.60 %\n\n\n", "Inspecci_n_Visual"=>[{"s1"=>"1. Cableado", "Cable_tipo_XHHW"=>"false", "Color_Naranja___L1"=>"false", "Color_Caf____L2"=>"false", "Color_Verde___Tierra"=>"true", "Distancia_m_nima"=>"true", "Se_uso_cinta_de_aislar"=>"false", "Se_uso_grasa"=>"false", "s2"=>"2. General", "La_terminaci_n_de_conectores_es_adecuada"=>"true", "Interruptores__soportes_y_conectores_fijos"=>"true", "Transformador_montado_adecuadamente"=>"true", "Recept_culo_de_Puesta_Tierra_y_Cables"=>"true", "s3"=>"3. Barra de Referencia a Tierra", "Conductor_para_alimentador_del_primario"=>"true", "Envolvente_del_panel"=>"true", "Pantalla_electrost_tica_del_transformador"=>"true", "Canalizaciones_de_los_derivados"=>"true", "Conductores_aislados_de_los_derivados"=>"true", "Monitor_de_aislamiento_de_linea"=>"true", "Cajas_y_envolventes_de_equipo_fijo"=>"true", "Observaciones"=>"No se usan los colores pero sin embargo hay identificación de colores rojo y blanco no es el tipo de cable utilizan THHW \n y canalización compartida", "Otras_observaciones"=>[{"Cable_desnudo"=>"false", "Recept_culos_con_terminal_de_puesta_a_tierra_aislada"=>"false", "Circuitos_derivados_tiene_mas_de_un_contacto_d_plex"=>"true", "Se_usan_cables_verdes_para_equipo_movil_"=>"false", "Cables_para_alarmas_remotas"=>"false", "Un_tablero_por_cama"=>"false", "UCI_minimo_16_receptaculos"=>"true", "Receptaculos_del_sistema_de_normal"=>"false"}]}], "Mediciones"=>[{"TR___FHC___THC___MHC"=>[{"Voltaje_Primario__medido_"=>"223.5", "Voltaje_Secundario__medido_"=>"126.3", "L1_vs_Tierra"=>"76.8", "L2_vs_Tierra"=>"32.832", "TR_L1____LIM_conectado"=>"49.9", "TR_L2___LIM_conectado"=>"33.9", "TR_L1___LIM_desconectado"=>"34.3", "TR_L2____LIM_desconectado"=>"32.4", "FHC_L1"=>"177.1", "FHC_L2"=>"208.4", "THC_L1"=>"173.4", "THC_L2"=>"223.9", "MHC_L1"=>"-3.7", "MHC_L2"=>"15.5", "__MHC_Cumple_Normatividad_"=>"No cumple"}], "__Interruptores_Derivados"=>"6", "__Circuitos_Disponibles_Sin_uso"=>"10", "_Circuitos_est_n_Identificados_"=>"false", "Circuitos"=>[{"Contacto"=>"Contacto 1 en sala 2", "L1"=>"78.9", "L2"=>"105", "Cumple__No_cumple"=>"Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Esta compartido a otra sala"}, {"Contacto"=>"Primero derecho ", "L1"=>"80", "L2"=>"104", "Cumple__No_cumple"=>"Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Primero derecha sala 2", "L1"=>"108", "L2"=>"142", "Cumple__No_cumple"=>"Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Compartido en sala 2"}, {"Contacto"=>"Tercer duplex", "L1"=>"600", "L2"=>"800", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Modulo", "L1"=>"44.3", "L2"=>"63.3", "Cumple__No_cumple"=>"Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Quinto duplex", "L1"=>"936", "L2"=>"885", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}]}], "Evidencia_Fotografica"=>[{"Puesta_en_marcha"=>[{"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Tablero "}, {"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Modulos"}, {"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Pared sala 1"}]}]}, {"_rea_a_revisar"=>"Terapia 12", "Existe_Codigo_QR_"=>"true", "Codigo_QR"=>"Cliente:  SOTO REYES ALEJANDRA\nNumero de Control:    1607217935\n\nDATOS DEL TABLERO\nModelo:  HIDP-2\t# de Serie:  48695\tWO:  220503\nDWG:   17053-A\tPanel:  HIDP2DB012CFNF01-DSU\nInt Gral :   15 Amp\tPolos:        2 \tMarca:   C.H.\n# de Ints deriv:   6             \tPolos:        2\tMarca:   C.H.\tAmps:  20\n\nDATOS DEL LIM\n# de Serie:  27523\tModelo:  MARK V\tDate Code:  0916\n\t\t\n\nDATOS DEL TRANSFORMADOR\nCapacidad:    2 KVA  \t# Catalogo:  21-0244L\nPrimario:  220 V\tSecundario:    120 V  \n# de Serie:  P162530136 \tImpedancia:   2.60 %\n", "Inspecci_n_Visual"=>[{"s1"=>"1. Cableado", "Cable_tipo_XHHW"=>"false", "Color_Naranja___L1"=>"false", "Color_Caf____L2"=>"false", "Color_Verde___Tierra"=>"true", "Distancia_m_nima"=>"true", "Se_uso_cinta_de_aislar"=>"false", "Se_uso_grasa"=>"false", "s2"=>"2. General", "La_terminaci_n_de_conectores_es_adecuada"=>"true", "Interruptores__soportes_y_conectores_fijos"=>"true", "Transformador_montado_adecuadamente"=>"true", "Recept_culo_de_Puesta_Tierra_y_Cables"=>"true", "s3"=>"3. Barra de Referencia a Tierra", "Conductor_para_alimentador_del_primario"=>"true", "Envolvente_del_panel"=>"true", "Pantalla_electrost_tica_del_transformador"=>"true", "Canalizaciones_de_los_derivados"=>"true", "Conductores_aislados_de_los_derivados"=>"true", "Monitor_de_aislamiento_de_linea"=>"true", "Cajas_y_envolventes_de_equipo_fijo"=>"true", "Observaciones"=>"Usan cable THHW y canalización compartida", "Otras_observaciones"=>[{"Cable_desnudo"=>"false", "Recept_culos_con_terminal_de_puesta_a_tierra_aislada"=>"false", "Circuitos_derivados_tiene_mas_de_un_contacto_d_plex"=>"true", "Se_usan_cables_verdes_para_equipo_movil_"=>"false", "Cables_para_alarmas_remotas"=>"false", "Un_tablero_por_cama"=>"false", "UCI_minimo_16_receptaculos"=>"true", "Receptaculos_del_sistema_de_normal"=>"false"}]}], "Mediciones"=>[{"TR___FHC___THC___MHC"=>[{"Voltaje_Primario__medido_"=>"220", "Voltaje_Secundario__medido_"=>"122.8", "L1_vs_Tierra"=>"73", "L2_vs_Tierra"=>"33", "TR_L1____LIM_conectado"=>"33", "TR_L2___LIM_conectado"=>"50", "TR_L1___LIM_desconectado"=>"34.1", "TR_L2____LIM_desconectado"=>"32.1", "FHC_L1"=>"186.3", "FHC_L2"=>"222.6", "THC_L1"=>"187", "THC_L2"=>"239", "MHC_L1"=>"0.7", "MHC_L2"=>"16.4", "__MHC_Cumple_Normatividad_"=>"No cumple"}], "__Interruptores_Derivados"=>"6", "__Circuitos_Disponibles_Sin_uso"=>"10", "_Circuitos_est_n_Identificados_"=>"false", "Circuitos"=>[{"Contacto"=>"Pastilla 7 quinto duplex ", "L1"=>"890", "L2"=>"840", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 8 contactos sala11", "L1"=>"114.8", "L2"=>"164", "Cumple__No_cumple"=>"Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Compartido sala 11"}, {"Contacto"=>"Pastilla 9 primer duplex derecha ", "L1"=>"882", "L2"=>"837", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 10 modulo sala 11", "L1"=>"80", "L2"=>"105", "Cumple__No_cumple"=>"Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 11 3 duplex derecha", "L1"=>"889", "L2"=>"837", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 12 modulo contactos sala 12", "L1"=>"835", "L2"=>"890", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true"}]}], "Evidencia_Fotografica"=>[{"Puesta_en_marcha"=>[{"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Pared médica sala 12"}, {"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Tablero sala 12"}]}]}, {"_rea_a_revisar"=>"Terapia 3", "Existe_Codigo_QR_"=>"true", "Codigo_QR"=>"Cliente:  SOTO REYES ALEJANDRA\nNumero de Control:    1607217935\n\nDATOS DEL TABLERO\nModelo:  HIDP-2\t# de Serie:  48691\tWO:  220503\nDWG:   17053-A\tPanel:  HIDP2DB012CFNF01-DSU\nInt Gral :   15 Amp\tPolos:        2 \tMarca:   C.H.\n# de Ints deriv:   6             \tPolos:        2\tMarca:   C.H.\tAmps:  20\n\nDATOS DEL LIM\n# de Serie:  27516\tModelo:  MARK V\tDate Code:  0916\n\t\t\n\nDATOS DEL TRANSFORMADOR\nCapacidad:    2 KVA  \t# Catalogo:  21-0244L\nPrimario:  220 V\tSecundario:    120 V  \n# de Serie:  P162530132   \tImpedancia:   2.60 %\n", "Inspecci_n_Visual"=>[{"s1"=>"1. Cableado", "Cable_tipo_XHHW"=>"false", "Color_Naranja___L1"=>"false", "Color_Caf____L2"=>"false", "Color_Verde___Tierra"=>"true", "Distancia_m_nima"=>"true", "Se_uso_cinta_de_aislar"=>"false", "Se_uso_grasa"=>"false", "s2"=>"2. General", "La_terminaci_n_de_conectores_es_adecuada"=>"true", "Interruptores__soportes_y_conectores_fijos"=>"true", "Transformador_montado_adecuadamente"=>"true", "Recept_culo_de_Puesta_Tierra_y_Cables"=>"true", "s3"=>"3. Barra de Referencia a Tierra", "Conductor_para_alimentador_del_primario"=>"true", "Envolvente_del_panel"=>"true", "Pantalla_electrost_tica_del_transformador"=>"true", "Canalizaciones_de_los_derivados"=>"true", "Conductores_aislados_de_los_derivados"=>"true", "Monitor_de_aislamiento_de_linea"=>"true", "Cajas_y_envolventes_de_equipo_fijo"=>"true", "Observaciones"=>"Usan cable THHW y canalización compartida ", "Otras_observaciones"=>[{"Cable_desnudo"=>"false", "Recept_culos_con_terminal_de_puesta_a_tierra_aislada"=>"false", "Circuitos_derivados_tiene_mas_de_un_contacto_d_plex"=>"true", "Se_usan_cables_verdes_para_equipo_movil_"=>"false", "Cables_para_alarmas_remotas"=>"false", "Un_tablero_por_cama"=>"false", "UCI_minimo_16_receptaculos"=>"true", "Receptaculos_del_sistema_de_normal"=>"false"}]}], "Mediciones"=>[{"TR___FHC___THC___MHC"=>[{"Voltaje_Primario__medido_"=>"221", "Voltaje_Secundario__medido_"=>"124", "L1_vs_Tierra"=>"69", "L2_vs_Tierra"=>"37", "TR_L1____LIM_conectado"=>"50", "TR_L2___LIM_conectado"=>"34", "TR_L1___LIM_desconectado"=>"34", "TR_L2____LIM_desconectado"=>"32", "FHC_L1"=>"190", "FHC_L2"=>"162", "THC_L1"=>"205", "THC_L2"=>"160", "MHC_L1"=>"15", "MHC_L2"=>"-2", "__MHC_Cumple_Normatividad_"=>"No cumple"}], "__Interruptores_Derivados"=>"6", "__Circuitos_Disponibles_Sin_uso"=>"10", "_Circuitos_est_n_Identificados_"=>"false", "Circuitos"=>[{"Contacto"=>"Pastilla 1 módulo sala 4", "L1"=>"99", "L2"=>"72", "Cumple__No_cumple"=>"Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Compartido a sala 4"}, {"Contacto"=>"Pastilla 2 duplex 1-2", "L1"=>"840", "L2"=>"880", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 3 contactos sala 4", "L1"=>"899", "L2"=>"844", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Esta en sala 4"}, {"Contacto"=>"Pastilla 4 duplex 2-4 derecha", "L1"=>"840", "L2"=>"888", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 5 modulo contactos sala 3", "L1"=>"890", "L2"=>"840", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 6 contacto 5 izquierda ", "L1"=>"840", "L2"=>"880", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}]}], "Evidencia_Fotografica"=>[{"Puesta_en_marcha"=>[{"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Pared médica sala 3"}, {"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Tablero sala 3"}]}]}, {"_rea_a_revisar"=>"Terapia 10", "Existe_Codigo_QR_"=>"true", "Codigo_QR"=>"Cliente:  SOTO REYES ALEJANDRA\nNumero de Control:  1607217935\n\nDATOS DEL TABLERO\nModelo:  HIDP-2  \t# de Serie: 48690\tWO:  220503\nDWG:  17053-A\tPanel:  HIDP2DB012CFNF01-DSU\nInt Gral:   15 Amp\tPolos:    2\tMarca:    C.H.\n# de Ints deriv:   6              \tPolos:    2\tMarca:    C.H.\tAmps:  20  \n\n\nDATOS DEL LIM\n# de Serie:  27518\tModelo:  MARK V\tDate Code:  0916\n\t\t\n\t\t\n\nDATOS DEL TRANSFORMADOR\nCapacidad:   2 KVA    \t# Catalogo:  21-0244L\nPrimario:    220 V   \tSecundario:    120 V\n# de Serie:  P162530131\t   Impedancia :   2.60 %\n", "Inspecci_n_Visual"=>[{"s1"=>"1. Cableado", "Cable_tipo_XHHW"=>"false", "Color_Naranja___L1"=>"false", "Color_Caf____L2"=>"false", "Color_Verde___Tierra"=>"true", "Distancia_m_nima"=>"true", "Se_uso_cinta_de_aislar"=>"false", "Se_uso_grasa"=>"false", "s2"=>"2. General", "La_terminaci_n_de_conectores_es_adecuada"=>"true", "Interruptores__soportes_y_conectores_fijos"=>"true", "Transformador_montado_adecuadamente"=>"true", "Recept_culo_de_Puesta_Tierra_y_Cables"=>"true", "s3"=>"3. Barra de Referencia a Tierra", "Conductor_para_alimentador_del_primario"=>"true", "Envolvente_del_panel"=>"true", "Pantalla_electrost_tica_del_transformador"=>"true", "Canalizaciones_de_los_derivados"=>"true", "Conductores_aislados_de_los_derivados"=>"true", "Monitor_de_aislamiento_de_linea"=>"true", "Cajas_y_envolventes_de_equipo_fijo"=>"true", "Observaciones"=>"Usan cable THHW y canalización compartida", "Otras_observaciones"=>[{"Cable_desnudo"=>"false", "Recept_culos_con_terminal_de_puesta_a_tierra_aislada"=>"false", "Circuitos_derivados_tiene_mas_de_un_contacto_d_plex"=>"true", "Se_usan_cables_verdes_para_equipo_movil_"=>"false", "Cables_para_alarmas_remotas"=>"false", "Un_tablero_por_cama"=>"false", "UCI_minimo_16_receptaculos"=>"true", "Receptaculos_del_sistema_de_normal"=>"false"}]}], "Mediciones"=>[{"TR___FHC___THC___MHC"=>[{"Voltaje_Primario__medido_"=>"218", "Voltaje_Secundario__medido_"=>"122", "L1_vs_Tierra"=>"73", "L2_vs_Tierra"=>"33", "TR_L1____LIM_conectado"=>"31.8", "TR_L2___LIM_conectado"=>"45", "TR_L1___LIM_desconectado"=>"32.6", "TR_L2____LIM_desconectado"=>"31.5", "FHC_L1"=>"171.7", "FHC_L2"=>"187.6", "THC_L1"=>"169", "THC_L2"=>"200", "MHC_L1"=>"-2.7", "MHC_L2"=>"12.4", "__MHC_Cumple_Normatividad_"=>"No cumple"}], "__Interruptores_Derivados"=>"6", "__Circuitos_Disponibles_Sin_uso"=>"10", "_Circuitos_est_n_Identificados_"=>"false", "Circuitos"=>[{"Contacto"=>"Pastilla 1 módulo contactos sala 9", "L1"=>"812", "L2"=>"862", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Compartido de sala 10-9"}, {"Contacto"=>"Pastilla 2 duplex 1-2 izquierda ", "L1"=>"867", "L2"=>"820", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 3 contactos duplex sala 9", "L1"=>"820", "L2"=>"870", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Compartido salas 10-9"}, {"Contacto"=>"Pastilla 4 duplex 3-4", "L1"=>"875", "L2"=>"827", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 5 modulo contactos ", "L1"=>"830", "L2"=>"880", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"false", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 6 duplex 5", "L1"=>"872", "L2"=>"830", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}]}], "Evidencia_Fotografica"=>[{"Puesta_en_marcha"=>[{"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Pared médica "}, {"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Tablero "}]}]}, {"_rea_a_revisar"=>"Terapia 8", "Existe_Codigo_QR_"=>"true", "Codigo_QR"=>"\nCliente:  SOTO REYES ALEJANDRA\nNumero de Control:    1607217935\n\nDATOS DEL TABLERO\nModelo:  HIDP-2\t# de Serie:  48693\tWO:  220503\nDWG:   17053-A\tPanel:  HIDP2DB012CFNF01-DSU\nInt Gral :   15 Amp\tPolos:        2 \tMarca:   C.H.\n# de Ints deriv:   6             \tPolos:        2\tMarca:   C.H.\tAmps:  20\n\nDATOS DEL LIM\n# de Serie:  27515\tModelo:  MARK V\tDate Code:  0916\n\t\t\n\nDATOS DEL TRANSFORMADOR\nCapacidad:    2 KVA  \t# Catalogo:  21-0244L\nPrimario:  220 V\tSecundario:    120 V  \n# de Serie:  P162530134  \tImpedancia:   2.60 %\n", "Inspecci_n_Visual"=>[{"s1"=>"1. Cableado", "Cable_tipo_XHHW"=>"false", "Color_Naranja___L1"=>"false", "Color_Caf____L2"=>"false", "Color_Verde___Tierra"=>"true", "Distancia_m_nima"=>"true", "Se_uso_cinta_de_aislar"=>"false", "Se_uso_grasa"=>"false", "s2"=>"2. General", "La_terminaci_n_de_conectores_es_adecuada"=>"true", "Interruptores__soportes_y_conectores_fijos"=>"true", "Transformador_montado_adecuadamente"=>"true", "Recept_culo_de_Puesta_Tierra_y_Cables"=>"true", "s3"=>"3. Barra de Referencia a Tierra", "Conductor_para_alimentador_del_primario"=>"true", "Envolvente_del_panel"=>"true", "Pantalla_electrost_tica_del_transformador"=>"true", "Canalizaciones_de_los_derivados"=>"true", "Conductores_aislados_de_los_derivados"=>"true", "Monitor_de_aislamiento_de_linea"=>"true", "Cajas_y_envolventes_de_equipo_fijo"=>"true", "Observaciones"=>"Usan cable THHW y canalización compartida", "Otras_observaciones"=>[{"Cable_desnudo"=>"false", "Recept_culos_con_terminal_de_puesta_a_tierra_aislada"=>"false", "Circuitos_derivados_tiene_mas_de_un_contacto_d_plex"=>"true", "Se_usan_cables_verdes_para_equipo_movil_"=>"false", "Cables_para_alarmas_remotas"=>"false", "Un_tablero_por_cama"=>"false", "UCI_minimo_16_receptaculos"=>"true", "Receptaculos_del_sistema_de_normal"=>"false"}]}], "Mediciones"=>[{"TR___FHC___THC___MHC"=>[{"Voltaje_Primario__medido_"=>"219", "Voltaje_Secundario__medido_"=>"123", "L1_vs_Tierra"=>"69.7", "L2_vs_Tierra"=>"37", "TR_L1____LIM_conectado"=>"34", "TR_L2___LIM_conectado"=>"48.8", "TR_L1___LIM_desconectado"=>"32.5", "TR_L2____LIM_desconectado"=>"31.4", "FHC_L1"=>"190", "FHC_L2"=>"170", "THC_L1"=>"190", "THC_L2"=>"185", "MHC_L1"=>"0", "MHC_L2"=>"15", "__MHC_Cumple_Normatividad_"=>"No cumple"}], "__Interruptores_Derivados"=>"6", "__Circuitos_Disponibles_Sin_uso"=>"10", "_Circuitos_est_n_Identificados_"=>"false", "Circuitos"=>[{"Contacto"=>"Pastilla 1 módulo contactos sala 7", "L1"=>"873", "L2"=>"820", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Compartido sala 8-7 "}, {"Contacto"=>"Pastilla 2 duplex 1-2", "L1"=>"865", "L2"=>"815", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 3 contactos duplex sala 7", "L1"=>"872", "L2"=>"820", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Compartido sala 8-7 "}, {"Contacto"=>"Pastilla 4 duplex 3-4", "L1"=>"857", "L2"=>"790", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 5 modulo de contactos ", "L1"=>"855", "L2"=>"800", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 6 duplex 5", "L1"=>"867", "L2"=>"815", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}]}], "Evidencia_Fotografica"=>[{"Puesta_en_marcha"=>[{"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Pared médica "}, {"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Tablero "}]}]}, {"_rea_a_revisar"=>"Terapia 5", "Existe_Codigo_QR_"=>"true", "Codigo_QR"=>"Cliente:  SOTO REYES ALEJANDRA\nNumero de Control:  1607217935\n\nDATOS DEL TABLERO\nModelo:  HIDP-2  \t# de Serie: 48692\tWO:  220503\nDWG:  17053-A\tPanel:  HIDP2DB012CFNF01-DSU\nInt Gral:   15 Amp\tPolos:    2\tMarca:    C.H.\n# de Ints deriv:   6              \tPolos:    2\tMarca:    C.H.\tAmps:  20  \n\n\nDATOS DEL LIM\n# de Serie:  27517\tModelo:  MARK V\tDate Code:  0916\n\t\t\n\t\t\n\nDATOS DEL TRANSFORMADOR\nCapacidad:   2 KVA    \t# Catalogo:  21-0244L\nPrimario:    220 V   \tSecundario:    120 V\n# de Serie:  P162530133\t    Impedancia :   2.60 %\n\n", "Inspecci_n_Visual"=>[{"s1"=>"1. Cableado", "Cable_tipo_XHHW"=>"false", "Color_Naranja___L1"=>"false", "Color_Caf____L2"=>"false", "Color_Verde___Tierra"=>"true", "Distancia_m_nima"=>"true", "Se_uso_cinta_de_aislar"=>"false", "Se_uso_grasa"=>"false", "s2"=>"2. General", "La_terminaci_n_de_conectores_es_adecuada"=>"true", "Interruptores__soportes_y_conectores_fijos"=>"true", "Transformador_montado_adecuadamente"=>"true", "Recept_culo_de_Puesta_Tierra_y_Cables"=>"true", "s3"=>"3. Barra de Referencia a Tierra", "Conductor_para_alimentador_del_primario"=>"true", "Envolvente_del_panel"=>"true", "Pantalla_electrost_tica_del_transformador"=>"true", "Canalizaciones_de_los_derivados"=>"true", "Conductores_aislados_de_los_derivados"=>"true", "Monitor_de_aislamiento_de_linea"=>"true", "Cajas_y_envolventes_de_equipo_fijo"=>"true", "Observaciones"=>"Usan cable THHW y canalización compartida", "Otras_observaciones"=>[{"Cable_desnudo"=>"false", "Recept_culos_con_terminal_de_puesta_a_tierra_aislada"=>"false", "Circuitos_derivados_tiene_mas_de_un_contacto_d_plex"=>"true", "Se_usan_cables_verdes_para_equipo_movil_"=>"false", "Cables_para_alarmas_remotas"=>"false", "Un_tablero_por_cama"=>"false", "UCI_minimo_16_receptaculos"=>"true", "Receptaculos_del_sistema_de_normal"=>"false"}]}], "Mediciones"=>[{"TR___FHC___THC___MHC"=>[{"Voltaje_Primario__medido_"=>"217", "Voltaje_Secundario__medido_"=>"121.8", "L1_vs_Tierra"=>"70", "L2_vs_Tierra"=>"32", "TR_L1____LIM_conectado"=>"32", "TR_L2___LIM_conectado"=>"47.5", "TR_L1___LIM_desconectado"=>"33", "TR_L2____LIM_desconectado"=>"32", "FHC_L1"=>"237", "FHC_L2"=>"200", "THC_L1"=>"233", "THC_L2"=>"215", "MHC_L1"=>"-4", "MHC_L2"=>"15", "__MHC_Cumple_Normatividad_"=>"No cumple"}], "__Interruptores_Derivados"=>"6", "__Circuitos_Disponibles_Sin_uso"=>"10", "_Circuitos_est_n_Identificados_"=>"false", "Circuitos"=>[{"Contacto"=>"Pastilla 1 módulo contactos sala 6", "L1"=>"870", "L2"=>"820", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Compartido sala 5-6"}, {"Contacto"=>"Pastilla 2 contactos 1-2 izquierda vertical ", "L1"=>"872", "L2"=>"825", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 3 contactos duplex sala 6", "L1"=>"873", "L2"=>"818", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 4 duplex 2-3 izquierda vertical ", "L1"=>"865", "L2"=>"815", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Alimentación de 4 duplex "}, {"Contacto"=>"Pastilla 5 modulo contactos ", "L1"=>"870", "L2"=>"822", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true"}, {"Contacto"=>"Pastilla 6 duplex 4-5 izquierda vertical ", "L1"=>"876", "L2"=>"816", "Cumple__No_cumple"=>"No Cumple", "Cumple_con_Polaridad"=>"true", "Cumple_con_Retenci_n"=>"true", "Comentarios"=>"Alimentación de 4 duplex "}]}], "Evidencia_Fotografica"=>[{"Puesta_en_marcha"=>[{"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Pared médica "}, {"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Tablero "}, {"Evidencia_Fotografica"=>"IMAGE", "Descripci_n"=>"Contacto "}]}]}], "Datos_Cliente"=>[{"Nombre_del_Cliente"=>"Constructora nortag sa de CV ", "Telefono_Cliente"=>"4771479933", "Email_del_Cliente"=>"constructoranortag@gmail.com", "Firma_del_Cliente"=>"IMAGE", "Observaciones_del_Cliente_"=>"Muy eficiente en pruebas y muy limpio "}], "Fecha_del_proximo_mantenimiento"=>"2017-10-28T09:12:31-0500", "Firma_del_T_cnico"=>"IMAGE", "Email_Vendedor"=>"erasmo.cruz@biors.com.mx", "Email_Personal_de_Servicio"=>"servicios@biors.com.mx"}
  #   }
  #   check_template(path_to_template, path_to_correct_render, {render_params: render_params})
  # end
end
