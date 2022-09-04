package com.rbard.jxlsexportexcel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@SpringBootApplication
public class JxlsExportExcelApplication {

    public static void main(String[] args) {
        SpringApplication.run(JxlsExportExcelApplication.class, args);
    }

    @GetMapping(value = "/generate-excel", produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public ResponseEntity<Resource> generateExcel() throws IOException {

        // obteniendo la plantilla
        byte[] byteInputTemplate = Files.readAllBytes(
                Path.of("src" + File.separator + "main" + File.separator + "resources"
                        + File.separator + "template.xlsx"));

        // obteniendo la data
        Map<String, Object> mapPerson = new HashMap<>();
        mapPerson.put("name", "joe");
        mapPerson.put("age", 30);
        mapPerson.put("email", "joe@gmail.com");

        List<Map<String, Object>> listReports = new ArrayList<>();
        Map<String, Object> mapReport1 = new HashMap<>();
        mapReport1.put("course", "mathematics");
        mapReport1.put("grade", 15.5);
        listReports.add(mapReport1);

        Map<String, Object> mapReport2 = new HashMap<>();
        mapReport2.put("course", "geography");
        mapReport2.put("grade", 17.0);
        listReports.add(mapReport2);

        Map<String, Object> mapReport3 = new HashMap<>();
        mapReport3.put("course", "english");
        mapReport3.put("grade", 18.5);
        listReports.add(mapReport3);

        Map<String, Object> mapData = new HashMap<>();
        mapData.put("person", mapPerson);
        mapData.put("reports", listReports);

        Context context = new Context();

        // Ingresando data en el contexto
        mapData.forEach((s, o) -> context.putVar(s, o));

        byte[] byteOutputTemplate;
        try (
                InputStream templateStream = new ByteArrayInputStream(byteInputTemplate);
                ByteArrayOutputStream targetStream = new ByteArrayOutputStream()
        ) {
            // Ingresando el contexto en la plantilla
            JxlsHelper helper = JxlsHelper.getInstance();
            Transformer transformer = PoiTransformer.createTransformer(templateStream,
                    targetStream);
            helper.processTemplate(context, transformer);
            byteOutputTemplate = targetStream.toByteArray();
        }

        InputStream inputStream = new ByteArrayInputStream(byteOutputTemplate);

        HttpHeaders httpHeaders = new HttpHeaders();
        httpHeaders.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=invoice.xlsx");
        httpHeaders.add(HttpHeaders.CONTENT_LENGTH, String.valueOf(byteOutputTemplate.length));

        return new ResponseEntity<>(new InputStreamResource(inputStream), httpHeaders,
                HttpStatus.OK);
    }

}
