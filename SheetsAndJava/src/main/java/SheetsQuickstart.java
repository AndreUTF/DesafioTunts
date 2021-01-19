import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class SheetsQuickstart {
    private static final String APPLICATION_NAME = "sheetsandjava";
    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private static final String TOKENS_DIRECTORY_PATH = "tokens";

    /**
     * Global instance of the scopes required by this quickstart.
     * If modifying these scopes, delete your previously saved tokens/ folder.
     */
    //Código referente a amostra disponível do Google Sheets API
    private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS);
    private static final String CREDENTIALS_FILE_PATH = "/credentials3.json";

    //Código referente a amostra disponível do Google Sheets API
    private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
        // Load client secrets.
        InputStream in = SheetsQuickstart.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
        if (in == null) {
            throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
        }
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        //Código referente a amostra disponível do Google Sheets API
        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
                HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
                .setAccessType("offline")
                .build();
        LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
        return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    }

    public static void main(String... args) throws IOException, GeneralSecurityException {
        //Código referente a amostra disponível do Google Sheets API
        // Build a new authorized API client service.
        final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();

        //ID referente a planilha initulada "Cópia de Engenharia de Software - Desafio André Portes"
        //Disponível em https://docs.google.com/spreadsheets/d/16roxR7LxF0mDmowjyTvCBO0tzoJhcYs0UIo1NrQWCC8/edit#gid=0
        final String spreadsheetId = "16roxR7LxF0mDmowjyTvCBO0tzoJhcYs0UIo1NrQWCC8";

        //Noma da planilha em que estão contidos os dados de interesse e intervalo de interese em que estão os dados das notas e frequência escolar
        final String range = "engenharia_de_software!A4:F27";

        //Código referente a amostra disponível do Google Sheets API
        Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
                .setApplicationName(APPLICATION_NAME)
                .build();

        //Código referente a amostra disponível do Google Sheets API
        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();
        //Código referente a amostra disponível do Google Sheets API
        List<List<Object>> values = response.getValues();

        //Array list de Strings onde serão armazenadas as informações da situção e nota do exame final de cada aluno
        ArrayList<String> arraySituacao = new ArrayList<String>();
        ArrayList<String> arrayNaf = new ArrayList<String>();

        //Código referente a amostra disponível do Google Sheets API
        if (values == null || values.isEmpty()) {
            System.out.println("No data found.");
        } else {
            //Loop que percorre as 24 linhas com as informações das notas e faltas dos alunos
            for (List row : values) {

                //Variáveis utilizadas para a determinação da situação do aluno, assim como a nota do exame final para aprovação
                float media = 0;
                float soma = 0;
                int faltas = 0;
                float naf_nota = 0;

                //Strings onde serão utilizadas para guardar a situação e nota final dos alunos
                String situacao = null;
                String naf = null;

                //Soma das notas de cada aluno utilizando o método Float.parserFloat a fim de conver texto em um float
                soma = Float.parseFloat((String) row.get(3)) + Float.parseFloat((String) row.get(4))
                        + Float.parseFloat((String) row.get(5));

                //Cálculo das médias em float
                media = soma / 3;

                //Obtenção do número de faltas em inteiros pelo método Integer.parseInt()
                faltas = Integer.parseInt((String) row.get(2));

                //Cáculo da nota para quem ficou de exma final, pela fórmula 5<=(m+naf)/2
                naf_nota = (float) 100.0 - media;

                //Verifica se o aluno excedeu o numero de faltas máxima (25% de 60 aulas = 15 faltas no máximo)
                if (faltas > 15) {
                    //Caso positivo o aluno está reprovado por falta e sua nota na prova final será zero
                    situacao = "Reprovado por Falta";
                    naf = "0";

                    //Adiciona as informações anteriores nos Arrays' List respectivo
                    arraySituacao.add(situacao);
                    arrayNaf.add(naf);
                } else {
                    //Caso negativo, ou seja, o estudante não excedeu o número de faltas
                    // Será verifico a nota para assim ter sua situação e nota do exame final atualizados

                    //Caso média for menor que 50 ( nota vai de o a 100), o aluno está reprovado e com nota zero no exame final
                    if (media < 50) {
                        situacao = "Reprovado por Nota";
                        naf = "0";

                        //Adiciona as informações anteriores nos Arrays' List respectivo
                        arraySituacao.add(situacao);
                        arrayNaf.add(naf);
                    }
                    //Caso média for maior ou igual a 50 e menor ou igual a 70 ( nota vai de o a 100), o aluno está exame final e com nota do exame final calculada
                    if (media >= 50 && media <= 70) {
                        situacao = "Exame Final";
                        naf = String.valueOf(naf_nota);

                        //Adiciona as informações anteriores nos Arrays' List respectivo
                        arraySituacao.add(situacao);
                        arrayNaf.add(naf);
                    }
                    //Caso média for maior que 70 ( nota vai de o a 100), o aluno está aprovado e com nota do exame zero
                    if (media > 70) {
                        situacao = "Aprovado";
                        naf = "0";

                        //Adiciona as informações anteriores nos Arrays' List respectivo
                        arraySituacao.add(situacao);
                        arrayNaf.add(naf);
                    }
                }
            }
        }

        //Cria-se dois ArraysList a fim de obter corretamente os campos que serão atualizados para as situações e notas de exame final de cada estudante
        ArrayList compoSituacao = new ArrayList<String>();
        ArrayList compoNaf = new ArrayList<String>();
        for (int j = 4; j <= 27; j++) {
            String letra1 = "G";
            String letra2 = "H";
            compoSituacao.add(letra1 + String.valueOf(j));
            compoNaf.add(letra2 + String.valueOf(j));
        }

        //Atualiza-se a tabela nas campos de Situação e Nota do exame final para todos os alunos
        for (int j = 0; j < 24; j++) {
            //Obtem-se a situação do aluno j, pelo método get(j) e este valor é adicionado ao value range
            ValueRange body = new ValueRange()
                    .setValues(Arrays.asList(
                            Arrays.asList(arraySituacao.get(j))
                    ));
            //Atualiza-se a tabela no campo indicado por String.valueOf(compoSituacao.get(j)), com o dado obtido pelo ValueRange
            UpdateValuesResponse result = service.spreadsheets()
                    .values().update(spreadsheetId, String.valueOf(compoSituacao.get(j)), body)
                    .setValueInputOption("RAW").execute();

            //Obtem-se a situação do aluno j, pelo método get(j) e este valor é adicionado ao value range
            ValueRange body1 = new ValueRange()
                    .setValues(Arrays.asList(
                            Arrays.asList(arrayNaf.get(j))
                    ));

            //Atualiza-se a tabela no campo indicado por String.valueOf(compoSituacao.get(j)), com o dado obtido pelo ValueRange
            UpdateValuesResponse result1 = service.spreadsheets()
                    .values().update(spreadsheetId, String.valueOf(compoNaf.get(j)), body1)
                    .setValueInputOption("RAW").execute();
        }
    }
}