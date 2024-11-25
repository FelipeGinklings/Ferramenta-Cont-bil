import pandas as pd
import numpy as np
from typing import List
import os
import errno
import os.path  # Adicione esta importação


class Conciliation:
    _input_path: List[str]
    _output_path: str

    def __init__(self):
        self._input_path = []
        self._output_path = ""
        self._data_frame = None
        self._different_hist = None
        self._initial_data_frame = None
        self._completed_paid = None
        self._incomplete_payment = None
        self._last_year_payments = None
        self._next_year = None
        self._similar_values_df = None

    def _load_and_process_data(self, file):
        # Carrega o arquivo CSV e seleciona as colunas necessárias
        self._data_frame = pd.read_csv(file, sep=";")
        self._data_frame = self._data_frame.loc[:, ["Valor", "Hist", "Complemento"]]

        # Filtra as linhas com valores de "Hist" diferentes de 133 e 20
        self._different_hist = self._data_frame[
            (self._data_frame["Hist"] != 133) & (self._data_frame["Hist"] != 20)
        ]

        # Extrai IDs dos campos "Complemento"
        self._data_frame["Id"] = self._data_frame["Complemento"].str.extract(
            r"(\d+)", expand=False
        )
        formatted_id = self._data_frame["Complemento"].str.extract(
            r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", expand=False
        )
        self._data_frame["Id"] = formatted_id.fillna(self._data_frame["Id"])
        self._data_frame["Id"] = self._data_frame["Id"].fillna(
            self._data_frame["Complemento"]
        )

        # Salva o estado do data_frame após a linha 33
        self._initial_data_frame = self._data_frame.copy()

        # Calcula o valor assinado com base no campo "Hist"
        self._data_frame["signed_value"] = self._data_frame.apply(
            lambda row: (
                row["Valor"] * -1
                if row["Hist"] == 20
                else row["Valor"] * 1 if row["Hist"] == 133 else row["Valor"]
            ),
            axis=1,
        )

    def _calculate_results(self):
        # Agrupa os dados pelo campo "Id" e calcula a soma dos valores assinados
        result = (
            self._data_frame.groupby("Id")["signed_value"]
            .sum()
            .reset_index(name="Resultado")
        )
        # Ajusta os resultados próximos de zero para zero
        result["Resultado"] = np.where(
            result["Resultado"].abs() < 1e-10, 0, result["Resultado"].round(10)
        )
        # Separa os resultados em três categorias
        self._completed_paid = result[result["Resultado"] == 0]
        self._last_year_payments = result[result["Resultado"] > 0].sort_values(
            by="Resultado"
        )
        self._incomplete_payment = result[result["Resultado"] < 0].sort_values(
            by="Resultado"
        )
        # Adiciona a nova categoria next_year
        self._similar_values_df = self._find_similar_values(
            self._last_year_payments, self._incomplete_payment
        )
        self._next_year = self._incomplete_payment[
            self._incomplete_payment["Resultado"].isin(
                self._data_frame[self._data_frame["Hist"] == 20]["signed_value"]
            )
            & ~self._incomplete_payment["Id"].isin(
                self._similar_values_df["Id Negativo"]
            )
        ].sort_values(by="Resultado")
        # Remove os valores de next_year de incomplete_payment
        self._incomplete_payment = self._incomplete_payment[
            ~self._incomplete_payment["Id"].isin(self._next_year["Id"])
        ]

    def _find_similar_values(self, last_year_payment, incomplete_payment):
        # Encontra valores semelhantes entre as listas de resultados positivos e negativos
        similar_ids = []
        for idx, row in last_year_payment.iterrows():
            opposite_row = incomplete_payment[
                incomplete_payment["Resultado"].abs() == abs(row["Resultado"])
            ]
            if not opposite_row.empty:
                similar_ids.append(
                    {
                        "Id Positivo": row["Id"],
                        "Resultado Positivo": row["Resultado"],
                        "Id Negativo": opposite_row.iloc[0]["Id"],
                        "Resultado Negativo": opposite_row.iloc[0]["Resultado"],
                    }
                )
        similar_values_df = pd.DataFrame(similar_ids).sort_values(
            by="Resultado Positivo"
        )
        return similar_values_df

    def _save_to_excel(self, input_file):
        # Reseta os índices dos DataFrames
        self._completed_paid.reset_index(drop=True, inplace=True)
        self._last_year_payments.reset_index(drop=True, inplace=True)
        self._incomplete_payment.reset_index(drop=True, inplace=True)
        self._different_hist.reset_index(drop=True, inplace=True)
        self._similar_values_df.reset_index(drop=True, inplace=True)
        self._initial_data_frame.reset_index(drop=True, inplace=True)
        # Adiciona flag para destacar os valores semelhantes
        self._last_year_payments["Highlight"] = self._last_year_payments["Id"].isin(
            self._similar_values_df["Id Positivo"]
        )
        self._incomplete_payment["Highlight"] = self._incomplete_payment["Id"].isin(
            self._similar_values_df["Id Negativo"]
        )

        # Ordena colocando os destacados no topo
        self._last_year_payments = self._last_year_payments.sort_values(
            by=["Highlight", "Resultado"], ascending=[False, True]
        ).drop(columns=["Highlight"])
        self._incomplete_payment = self._incomplete_payment.sort_values(
            by=["Highlight", "Resultado"], ascending=[False, False]
        ).drop(columns=["Highlight"])

        # Garante que o diretório de saída exista
        os.makedirs(self._output_path, exist_ok=True)

        # Extrai o nome do arquivo de entrada sem a extensão
        input_filename = os.path.splitext(os.path.basename(input_file))[0]
        output_file = f"{self._output_path}/{input_filename}.xlsx"

        # Verifica se o arquivo de saída está acessível
        if os.path.exists(output_file):
            try:
                os.rename(output_file, output_file)
            except OSError as e:
                if e.errno == errno.EACCES:
                    raise PermissionError(f"Permission denied: '{output_file}'")

        self._initial_data_frame = self._initial_data_frame.sort_values(by=["Id"]).drop(
            columns=["Complemento"]
        )
        new_order = ["Id", "Hist", "Valor"]
        self._initial_data_frame = self._initial_data_frame[new_order]
        # Salva os resultados em um arquivo Excel
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            self._last_year_payments.to_excel(
                writer, sheet_name="Ano Passado", index=False
            )
            self._incomplete_payment.to_excel(
                writer, sheet_name="Pagamento Incompleto", index=False
            )
            self._next_year.to_excel(writer, sheet_name="Próximo Ano", index=False)
            self._different_hist.to_excel(
                writer, sheet_name="Hist Diferente de 20 e 133", index=False
            )
            self._completed_paid.to_excel(
                writer, sheet_name="Pagamento Completo", index=False
            )
            self._initial_data_frame.to_excel(
                writer, sheet_name="Planilha Limpa", index=False
            )

            workbook = writer.book
            font_format = workbook.add_format({"font_size": 14})
            bold_format = workbook.add_format({"font_size": 14, "bold": True})

            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                worksheet.set_column(
                    "A:Z", None, font_format
                )  # Define o tamanho da fonte para as células
                # Determine o número de colunas baseado no DataFrame
                sheet_values = self._last_year_payments.columns.values
                if sheet_name == "Hist Diferente de 20 e 133":
                    sheet_values = self._different_hist.columns.values
                elif sheet_name == "Planilha Limpa":
                    sheet_values = self._initial_data_frame.columns.values
                for col_num, value in enumerate(sheet_values):
                    worksheet.write(
                        0, col_num, value, bold_format
                    )  # Cabeçalhos com negrito e tamanho definido

            # Adiciona negrito para IDs semelhantes em "Valores Semelhantes"
            worksheet_last_year_payments = writer.sheets["Ano Passado"]
            worksheet_incomplete_payment = writer.sheets["Pagamento Incompleto"]
            for idx, row in self._similar_values_df.iterrows():
                if row["Id Positivo"] in self._last_year_payments["Id"].values:
                    worksheet_last_year_payments.set_row(idx + 1, None, bold_format)
                    worksheet_incomplete_payment.set_row(idx + 1, None, bold_format)

    def _result(self, file):
        # Executa o processo de conciliação para um arquivo
        self._load_and_process_data(file)
        self._calculate_results()
        self._save_to_excel(file)

    def set_output(self, new_directory: str):
        """
        Define o diretório de saída para os arquivos de resultado.

        Parâmetros:
        new_directory (str): O caminho do novo diretório de saída.
        """
        self._output_path = new_directory

    def new_conciliation(self, input_path: List[str]):
        """
        Inicia um novo processo de conciliação para os arquivos fornecidos.

        Parâmetros:
        input_path (List[str]): Lista de caminhos dos arquivos de entrada.
        """
        for file in input_path:
            self._result(file)


# Exemplo de uso da classe Conciliation
if __name__ == "__main__":
    # Cria uma instância da classe Conciliation
    conciliation = Conciliation()

    # Define o diretório de saída
    conciliation.set_output("./result")

    # Lista de arquivos de entrada
    arquivos_de_entrada = [
        "csv/original.csv",
    ]

    # Inicia o processo de conciliação
    conciliation.new_conciliation(arquivos_de_entrada)
