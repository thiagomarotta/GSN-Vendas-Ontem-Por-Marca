function capitalizeName(name) {
  if (!name || typeof name !== "string") return "";

  return name
    .toLowerCase()
    .split(/(\s|-|')/g)  // Mantém os separadores (espaço, hífen e apóstrofo) durante o split
    .map(part => {
      if (part.match(/[\s\-']/)) {
        return part;  // Mantém os separadores como estão
      }
      // Capitaliza a primeira letra de cada parte
      return part.charAt(0).toLocaleUpperCase("pt-BR") + part.slice(1);
    })
    .join("");
}

function formatCpfCnpj(numero) {
  if (!numero) return "";

  // Remove qualquer caractere não numérico
  const digits = numero.toString().replace(/\D/g, "");

  if (digits.length === 11) {
    // CPF
    return digits.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
  } else if (digits.length === 10) {
    // CPF com 10 dígitos (zero omitido no início)
    const padded = digits.padStart(11, "0");
    return padded.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
  } else if (digits.length === 14) {
    // CNPJ
    return digits.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, "$1.$2.$3/$4-$5");
  } else if (digits.length === 13) {
    // CNPJ com 13 dígitos (zero omitido no início)
    const padded = digits.padStart(14, "0");
    return padded.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, "$1.$2.$3/$4-$5");
  } else {
    // Não é um CPF ou CNPJ válido
    return numero;  // Retorna como veio (evita apagar dado errado)
  }
}
