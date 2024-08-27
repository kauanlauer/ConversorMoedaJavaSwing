namespace Conversor_de_Moedas_C_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            txtValor.Clear();
            lblResultado.Text = string.Empty;
            cmbOrigem.SelectedIndex = 0;
            cmbDestino.SelectedIndex = 1;
        }


        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnConverter_Click(object sender, EventArgs e)
        {
            // Verificando se o valor digitado é válido
            if (decimal.TryParse(txtValor.Text, out decimal valor))
            {
                string moedaOrigem = cmbOrigem.SelectedItem.ToString();
                string moedaDestino = cmbDestino.SelectedItem.ToString();

                decimal taxa = 0;

                // Definindo as taxas de câmbio (valores fictícios)
                if (moedaOrigem == "USD" && moedaDestino == "EUR")
                    taxa = 0.85m; // Exemplo: 1 USD = 0.85 EUR
                else if (moedaOrigem == "USD" && moedaDestino == "BRL")
                    taxa = 5.20m; // Exemplo: 1 USD = 5.20 BRL
                else if (moedaOrigem == "EUR" && moedaDestino == "USD")
                    taxa = 1.18m; // Exemplo: 1 EUR = 1.18 USD
                else if (moedaOrigem == "EUR" && moedaDestino == "BRL")
                    taxa = 6.10m; // Exemplo: 1 EUR = 6.10 BRL
                else if (moedaOrigem == "BRL" && moedaDestino == "USD")
                    taxa = 0.19m; // Exemplo: 1 BRL = 0.19 USD
                else if (moedaOrigem == "BRL" && moedaDestino == "EUR")
                    taxa = 0.16m; // Exemplo: 1 BRL = 0.16 EUR
                else
                    taxa = 1; // Caso as moedas sejam iguais

                // Realizando a conversão
                decimal resultado = valor * taxa;
                lblResultado.Text = $"Valor Convertido: {resultado:F2} {moedaDestino}";
            }
            else
            {
                MessageBox.Show("Por favor, insira um valor válido.");
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // Adicionando opções de moedas nos ComboBoxes
            cmbOrigem.Items.Add("USD");
            cmbOrigem.Items.Add("EUR");
            cmbOrigem.Items.Add("BRL");

            cmbDestino.Items.Add("USD");
            cmbDestino.Items.Add("EUR");
            cmbDestino.Items.Add("BRL");

            // Definindo as moedas padrão nos ComboBoxes
            cmbOrigem.SelectedIndex = 0;
            cmbDestino.SelectedIndex = 1;
        }


    }
}

