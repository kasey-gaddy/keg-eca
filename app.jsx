const { useState, useEffect, useCallback, useMemo } = React;
const XLSX = window.XLSX;

// ─── KE&G Brand ───────────────────────────────────────────────────────
const COLORS = {
  navy: "#1A2744",
  navyDark: "#0D1F3C",
  blue: "#103F91",
  blueMid: "#1B4F9B",
  orange: "#E8621A",
  orangeDark: "#E05A1A",
  white: "#FFFFFF",
  offWhite: "#F5F6F8",
  gray100: "#EEF0F4",
  gray200: "#D6DAE6",
  gray400: "#8C93A8",
  gray600: "#4B5270",
  green: "#1A7A4A",
  greenLight: "#E6F4EE",
  red: "#C0392B",
  redLight: "#FDEEEC",
  yellow: "#D4850A",
  yellowLight: "#FEF6E4",
};

const OFFICES = ["Tucson", "Sierra Vista", "Douglas (Maddux)"];

// Official 2026 roster: 610 active KE&G employees + 39 Maddux & Sons employees = 649 total.
// Loaded via Admin → Employees → "Load Official Roster" — never auto-applied, so it never
// silently overwrites live data. Update this array and re-load for future years.
const OFFICIAL_ROSTER = [
  { id: "emp0", number: "2209", firstName: "Jesus", lastName: "Abril", office: "Tucson", position: "Pipelayer" },
  { id: "emp1", number: "2131", firstName: "Adalberto", lastName: "Acuna", office: "Sierra Vista", position: "Laborer" },
  { id: "emp2", number: "26051", firstName: "Cesar", lastName: "Acuna", office: "Sierra Vista", position: "Superintendent" },
  { id: "emp3", number: "2946", firstName: "Michael", lastName: "Addison", office: "Sierra Vista", position: "Laborer" },
  { id: "emp4", number: "2802", firstName: "Victor", lastName: "Aguirre", office: "Tucson", position: "Pipelayer" },
  { id: "emp5", number: "1532", firstName: "Erik", lastName: "Ahkeah", office: "Tucson", position: "Foreman" },
  { id: "emp6", number: "1534", firstName: "Samuel", lastName: "Ahkeah", office: "Tucson", position: "Leadman" },
  { id: "emp7", number: "1791", firstName: "Manuel", lastName: "Ahumada", office: "Tucson", position: "Operator" },
  { id: "emp8", number: "2766", firstName: "Christopher", lastName: "Albright", office: "Tucson", position: "President" },
  { id: "emp9", number: "1109", firstName: "Kenneth", lastName: "Almodobar", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp10", number: "2652", firstName: "Tyler", lastName: "Alt", office: "Tucson", position: "Operator" },
  { id: "emp11", number: "3046", firstName: "Daniel", lastName: "Alvarado", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp12", number: "2213", firstName: "Adam", lastName: "Alvarez", office: "Tucson", position: "Operator" },
  { id: "emp13", number: "3134", firstName: "Alexander", lastName: "Alvarez", office: "Sierra Vista", position: "Laborer" },
  { id: "emp14", number: "3104", firstName: "Jason", lastName: "Alvarez", office: "Sierra Vista", position: "Project Engineer" },
  { id: "emp15", number: "1474", firstName: "Santana", lastName: "Alvarez", office: "Tucson", position: "Operator" },
  { id: "emp16", number: "2094", firstName: "Noe", lastName: "Amarillas", office: "Tucson", position: "Operator" },
  { id: "emp17", number: "1195", firstName: "Shane", lastName: "Amarillas", office: "Tucson", position: "Leadman" },
  { id: "emp18", number: "2502", firstName: "Justin", lastName: "Amick", office: "Tucson", position: "Operator" },
  { id: "emp19", number: "2120", firstName: "Kelly", lastName: "Amstutz", office: "Tucson", position: "Project Administrator" },
  { id: "emp20", number: "2897", firstName: "Cyruss", lastName: "Anderson", office: "Tucson", position: "Laborer" },
  { id: "emp21", number: "3110", firstName: "Dustin", lastName: "Anderson", office: "Sierra Vista", position: "Leadman" },
  { id: "emp22", number: "65064", firstName: "Wayne", lastName: "Anderson", office: "Tucson", position: "Vice President" },
  { id: "emp23", number: "1497", firstName: "Patrick", lastName: "Andrew", office: "Tucson", position: "Pipelayer" },
  { id: "emp24", number: "3049", firstName: "Bryan", lastName: "Arce-Marquez", office: "Tucson", position: "Laborer" },
  { id: "emp25", number: "3057", firstName: "Manuel", lastName: "Arellano", office: "Tucson", position: "Laborer" },
  { id: "emp26", number: "70793", firstName: "Ramon", lastName: "Armenta Jr", office: "Sierra Vista", position: "Foreman" },
  { id: "emp27", number: "2815", firstName: "Fabian", lastName: "Armenta Saenz", office: "Tucson", position: "Laborer" },
  { id: "emp28", number: "2974", firstName: "Victor", lastName: "Arreola", office: "Tucson", position: "Operator" },
  { id: "emp29", number: "1486", firstName: "Gabriel", lastName: "Arrieta", office: "Tucson", position: "Carpenter" },
  { id: "emp30", number: "2431", firstName: "Genavieve", lastName: "Arrieta", office: "Tucson", position: "" },
  { id: "emp31", number: "2827", firstName: "Carlos", lastName: "Arvayo", office: "Tucson", position: "Laborer" },
  { id: "emp32", number: "2312", firstName: "Armando", lastName: "Arvizu", office: "Tucson", position: "Operator" },
  { id: "emp33", number: "2980", firstName: "Truman", lastName: "Attakai", office: "Tucson", position: "Leadman" },
  { id: "emp34", number: "2941", firstName: "Ramon", lastName: "Avila", office: "Sierra Vista", position: "Pipelayer" },
  { id: "emp35", number: "3136", firstName: "Gabriel", lastName: "Aviles", office: "Sierra Vista", position: "Mechanic" },
  { id: "emp36", number: "1200", firstName: "Ivan", lastName: "Ayon", office: "Tucson", position: "Superintendent" },
  { id: "emp37", number: "2562", firstName: "Luis", lastName: "Ayon", office: "Tucson", position: "Laborer" },
  { id: "emp38", number: "2594", firstName: "James", lastName: "Baca", office: "Sierra Vista", position: "Operator" },
  { id: "emp39", number: "2929", firstName: "Cody", lastName: "Balke", office: "Sierra Vista", position: "Operator" },
  { id: "emp40", number: "2452", firstName: "Hector", lastName: "Ballesteros", office: "Sierra Vista", position: "Operator" },
  { id: "emp41", number: "1255", firstName: "Chris", lastName: "Baquera", office: "Tucson", position: "Superintendent" },
  { id: "emp42", number: "2722", firstName: "Laurence", lastName: "Barcelo Gomez", office: "Tucson", position: "Laborer" },
  { id: "emp43", number: "2021", firstName: "Robert", lastName: "Barney", office: "Tucson", position: "Operator" },
  { id: "emp44", number: "6850", firstName: "Jesus", lastName: "Barrera Valenzuela", office: "Sierra Vista", position: "Operator" },
  { id: "emp45", number: "2663", firstName: "Jesus", lastName: "Basurto Sesma", office: "Tucson", position: "Pipelayer" },
  { id: "emp46", number: "3042", firstName: "Hector", lastName: "Batriz", office: "Tucson", position: "Project Administrator" },
  { id: "emp47", number: "2818", firstName: "Russell", lastName: "Bauer-Woodman", office: "Tucson", position: "Field Engineer" },
  { id: "emp48", number: "2137", firstName: "Weslyn", lastName: "Bejarano", office: "Tucson", position: "" },
  { id: "emp49", number: "2881", firstName: "Lesley", lastName: "Bellai", office: "Tucson", position: "Vice President" },
  { id: "emp50", number: "2468", firstName: "Jason", lastName: "Bellomi", office: "Sierra Vista", position: "Project Engineer" },
  { id: "emp51", number: "2593", firstName: "Robert", lastName: "Bender", office: "Sierra Vista", position: "Laborer" },
  { id: "emp52", number: "2758", firstName: "Dylan", lastName: "Bernal", office: "Sierra Vista", position: "Laborer" },
  { id: "emp53", number: "79688", firstName: "Genaro", lastName: "Bernal", office: "Sierra Vista", position: "Operator" },
  { id: "emp54", number: "2133", firstName: "Francisco", lastName: "Betancourt", office: "Tucson", position: "Leadman" },
  { id: "emp55", number: "1915", firstName: "Javier", lastName: "Betancourt", office: "Tucson", position: "Leadman" },
  { id: "emp56", number: "2180", firstName: "Patrick", lastName: "Borgman", office: "Tucson", position: "Operator" },
  { id: "emp57", number: "2520", firstName: "Christopher", lastName: "Bowden", office: "Sierra Vista", position: "Foreman" },
  { id: "emp58", number: "2384", firstName: "Dylan", lastName: "Boyett", office: "Tucson", position: "" },
  { id: "emp59", number: "2243", firstName: "Jose", lastName: "Bracamontes", office: "Tucson", position: "Operator" },
  { id: "emp60", number: "2762", firstName: "Joel", lastName: "Brandau", office: "Sierra Vista", position: "Foreman" },
  { id: "emp61", number: "1422", firstName: "Gregorio", lastName: "Bray", office: "Tucson", position: "Superintendent" },
  { id: "emp62", number: "3010", firstName: "William", lastName: "Bridge", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp63", number: "2130", firstName: "Dakota", lastName: "Broughton", office: "Tucson", position: "" },
  { id: "emp64", number: "1342", firstName: "John", lastName: "Broughton", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp65", number: "2660", firstName: "Austin", lastName: "Brown", office: "Tucson", position: "Operator" },
  { id: "emp66", number: "2873", firstName: "Robert", lastName: "Brown", office: "Tucson", position: "Operator" },
  { id: "emp67", number: "2477", firstName: "Rustee", lastName: "Brown", office: "Tucson", position: "Laborer" },
  { id: "emp68", number: "2237", firstName: "Keith", lastName: "Buchanan", office: "Tucson", position: "Project Engineer" },
  { id: "emp69", number: "2778", firstName: "Cody", lastName: "Burke", office: "Tucson", position: "Foreman" },
  { id: "emp70", number: "33169", firstName: "Kenneth", lastName: "Burriss", office: "Tucson", position: "Foreman" },
  { id: "emp71", number: "2937", firstName: "Luis", lastName: "Bustamante Leon", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp72", number: "3015", firstName: "John", lastName: "Butler", office: "Tucson", position: "Operator" },
  { id: "emp73", number: "1599", firstName: "Daniel", lastName: "Buttram", office: "Tucson", position: "Superintendent" },
  { id: "emp74", number: "2824", firstName: "Anthony", lastName: "Caccavale", office: "Tucson", position: "Field Engineer" },
  { id: "emp75", number: "2532", firstName: "Jose", lastName: "Calderon", office: "Tucson", position: "Mechanic" },
  { id: "emp76", number: "96382", firstName: "Mark", lastName: "Campbell", office: "Tucson", position: "Division Manager" },
  { id: "emp77", number: "2341", firstName: "Armando", lastName: "Canez", office: "Tucson", position: "" },
  { id: "emp78", number: "2950", firstName: "Marcus", lastName: "Cannady Jr", office: "Sierra Vista", position: "Laborer" },
  { id: "emp79", number: "2859", firstName: "Omar", lastName: "Cantu", office: "Tucson", position: "Laborer" },
  { id: "emp80", number: "3008", firstName: "Roberto", lastName: "Cantua", office: "Tucson", position: "Laborer" },
  { id: "emp81", number: "1675", firstName: "Nicholas", lastName: "Carbajal", office: "Tucson", position: "Operator" },
  { id: "emp82", number: "2459", firstName: "Abraham", lastName: "Cardenas", office: "Tucson", position: "Laborer" },
  { id: "emp83", number: "2992", firstName: "Dabi", lastName: "Cardoza", office: "Tucson", position: "Operator" },
  { id: "emp84", number: "2880", firstName: "Victor", lastName: "Carey", office: "Tucson", position: "Operator" },
  { id: "emp85", number: "2726", firstName: "Jeffrey", lastName: "Carrico", office: "Tucson", position: "Foreman" },
  { id: "emp86", number: "2748", firstName: "Araceli", lastName: "Carrillo", office: "Tucson", position: "Project Administrator" },
  { id: "emp87", number: "3797", firstName: "Benjamin", lastName: "Carter", office: "Sierra Vista", position: "Project Manager_Estimator" },
  { id: "emp88", number: "2034", firstName: "Austin", lastName: "Carver", office: "Tucson", position: "Shop Laborer" },
  { id: "emp89", number: "2785", firstName: "Antonio", lastName: "Casillas", office: "Tucson", position: "Laborer" },
  { id: "emp90", number: "1771", firstName: "Arturo", lastName: "Castillo", office: "Tucson", position: "Foreman" },
  { id: "emp91", number: "2903", firstName: "Cesar", lastName: "Castillo", office: "Tucson", position: "Laborer" },
  { id: "emp92", number: "2964", firstName: "Jaime", lastName: "Castillo", office: "Tucson", position: "Operator" },
  { id: "emp93", number: "2501", firstName: "Mario", lastName: "Castillo", office: "Tucson", position: "Mechanic" },
  { id: "emp94", number: "2319", firstName: "Victor", lastName: "Castillo", office: "Sierra Vista", position: "" },
  { id: "emp95", number: "1727", firstName: "Phillip", lastName: "Castrillo", office: "Tucson", position: "Leadman" },
  { id: "emp96", number: "2984", firstName: "Cristian", lastName: "Castro", office: "Tucson", position: "Laborer" },
  { id: "emp97", number: "1618", firstName: "Roberto", lastName: "Castro", office: "Tucson", position: "Operator" },
  { id: "emp98", number: "2779", firstName: "Brian", lastName: "Cervantes", office: "Tucson", position: "Laborer" },
  { id: "emp99", number: "3056", firstName: "Abran", lastName: "Chacon Jr", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp100", number: "1962", firstName: "Francisco", lastName: "Chavarria", office: "Tucson", position: "Laborer" },
  { id: "emp101", number: "2456", firstName: "Miguel", lastName: "Chavez", office: "Tucson", position: "Pipelayer" },
  { id: "emp102", number: "2400", firstName: "Julio", lastName: "Chavez Gonzalez", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp103", number: "1418", firstName: "Wyatt", lastName: "Coates", office: "Tucson", position: "Operator" },
  { id: "emp104", number: "2895", firstName: "Isaiah", lastName: "Colmenero", office: "Tucson", position: "Laborer" },
  { id: "emp105", number: "2338", firstName: "Joel", lastName: "Colmenero", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp106", number: "2103", firstName: "Aaron", lastName: "Contreras", office: "Tucson", position: "Leadman" },
  { id: "emp107", number: "2495", firstName: "Angel", lastName: "Contreras Davis", office: "Tucson", position: "Pipelayer" },
  { id: "emp108", number: "3017", firstName: "Carlos", lastName: "Contreras Rojas", office: "Tucson", position: "Pipelayer" },
  { id: "emp109", number: "1189", firstName: "Adrian", lastName: "Cordova", office: "Tucson", position: "Operator" },
  { id: "emp110", number: "2888", firstName: "Andrew", lastName: "Cordova", office: "Tucson", position: "Laborer" },
  { id: "emp111", number: "2985", firstName: "Jesus", lastName: "Cordova", office: "Tucson", position: "Laborer" },
  { id: "emp112", number: "2887", firstName: "Manuel", lastName: "Cordova", office: "Tucson", position: "Pipelayer" },
  { id: "emp113", number: "1908", firstName: "Thomas", lastName: "Cordova", office: "Sierra Vista", position: "Operator" },
  { id: "emp114", number: "2460", firstName: "Marco", lastName: "Cordova Sevilla", office: "Tucson", position: "Laborer" },
  { id: "emp115", number: "8819", firstName: "Rodolfo", lastName: "Cordova-Encinas", office: "Tucson", position: "Foreman" },
  { id: "emp116", number: "2947", firstName: "Terrel", lastName: "Cork", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp117", number: "87196", firstName: "Vilionet", lastName: "Coronado", office: "Sierra Vista", position: "" },
  { id: "emp118", number: "2408", firstName: "Joel", lastName: "Coronado Acedo", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp119", number: "1667", firstName: "Perceo", lastName: "Coronado Quiroz", office: "Tucson", position: "Foreman" },
  { id: "emp120", number: "3014", firstName: "Reyes", lastName: "Corral", office: "Tucson", position: "Laborer" },
  { id: "emp121", number: "2848", firstName: "Logan", lastName: "Cosby", office: "Sierra Vista", position: "Operator" },
  { id: "emp122", number: "74247", firstName: "Daniel", lastName: "Crater", office: "Sierra Vista", position: "" },
  { id: "emp123", number: "2813", firstName: "Daniel", lastName: "Cross", office: "Tucson", position: "Laborer" },
  { id: "emp124", number: "3047", firstName: "David", lastName: "Cross", office: "Tucson", position: "Mechanic" },
  { id: "emp125", number: "21374", firstName: "Jayson", lastName: "Crunk", office: "Sierra Vista", position: "Leadman" },
  { id: "emp126", number: "3113", firstName: "Abelardo", lastName: "Cruz", office: "Sierra Vista", position: "Operator" },
  { id: "emp127", number: "1568", firstName: "Richard", lastName: "Cruz", office: "Tucson", position: "Mechanic" },
  { id: "emp128", number: "2904", firstName: "Tommy", lastName: "Cruz", office: "Tucson", position: "Operator" },
  { id: "emp129", number: "1926", firstName: "Fernando", lastName: "Cruz Munoz", office: "Tucson", position: "Operator" },
  { id: "emp130", number: "2569", firstName: "Georgina", lastName: "Cruz Olivares", office: "Tucson", position: "" },
  { id: "emp131", number: "1688", firstName: "Ceferino", lastName: "Cuevas Jr", office: "Tucson", position: "Pipelayer" },
  { id: "emp132", number: "2351", firstName: "Susana", lastName: "Cutuli", office: "Tucson", position: "Field Engineer" },
  { id: "emp133", number: "2165", firstName: "Jeffery", lastName: "Dahlen", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp134", number: "2629", firstName: "Eric", lastName: "Daily", office: "Tucson", position: "Safety Assistant" },
  { id: "emp135", number: "2412", firstName: "Gerardo", lastName: "De La Rosa", office: "Sierra Vista", position: "Mechanic" },
  { id: "emp136", number: "2858", firstName: "Jeramiah", lastName: "Delmar", office: "Tucson", position: "Laborer" },
  { id: "emp137", number: "2913", firstName: "Jacob", lastName: "Dewey", office: "Tucson", position: "" },
  { id: "emp138", number: "2977", firstName: "Marco", lastName: "Diaz Jr", office: "Tucson", position: "Operator" },
  { id: "emp139", number: "2631", firstName: "Jesus", lastName: "Dicochea", office: "Tucson", position: "Operator" },
  { id: "emp140", number: "2723", firstName: "Gilberto", lastName: "Dominguez", office: "Tucson", position: "Pipelayer" },
  { id: "emp141", number: "1943", firstName: "Martin", lastName: "Dominguez Garcia", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp142", number: "2936", firstName: "Roberto", lastName: "Dominguez Peralta", office: "Sierra Vista", position: "Operator" },
  { id: "emp143", number: "1999", firstName: "Jose", lastName: "Dorame", office: "Tucson", position: "Cdl Driver" },
  { id: "emp144", number: "14053", firstName: "John", lastName: "Drake", office: "Sierra Vista", position: "" },
  { id: "emp145", number: "3016", firstName: "Tommy", lastName: "Drew", office: "Tucson", position: "Pipelayer" },
  { id: "emp146", number: "1452", firstName: "Martin", lastName: "Duarte", office: "Tucson", position: "Leadman" },
  { id: "emp147", number: "1974", firstName: "Angel", lastName: "Duarte Duarte", office: "Tucson", position: "Laborer" },
  { id: "emp148", number: "2817", firstName: "Esdrel", lastName: "Duarte Favela", office: "Tucson", position: "Laborer" },
  { id: "emp149", number: "2650", firstName: "Robert", lastName: "Duchene", office: "Tucson", position: "Operator" },
  { id: "emp150", number: "2541", firstName: "Levi", lastName: "Duran", office: "Tucson", position: "Field Engineer" },
  { id: "emp151", number: "1179", firstName: "Ernest", lastName: "Durant", office: "Tucson", position: "Shop Laborer" },
  { id: "emp152", number: "2521", firstName: "Henoch", lastName: "Durazo", office: "Sierra Vista", position: "Operator" },
  { id: "emp153", number: "2845", firstName: "Noah", lastName: "Durazo", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp154", number: "31805", firstName: "Ricardo", lastName: "Durazo", office: "Sierra Vista", position: "Operator" },
  { id: "emp155", number: "1164", firstName: "Devon", lastName: "Eason", office: "Tucson", position: "Laborer" },
  { id: "emp156", number: "2794", firstName: "Matthew", lastName: "Edmiston", office: "Tucson", position: "Operator" },
  { id: "emp157", number: "2924", firstName: "Damian", lastName: "Elias", office: "Sierra Vista", position: "Laborer" },
  { id: "emp158", number: "2637", firstName: "Teresa", lastName: "Ely", office: "Tucson", position: "Operator" },
  { id: "emp159", number: "2867", firstName: "Benjamin", lastName: "Encinas", office: "Tucson", position: "Pipelayer" },
  { id: "emp160", number: "3116", firstName: "Judas", lastName: "Encinas", office: "Sierra Vista", position: "Laborer" },
  { id: "emp161", number: "2968", firstName: "Adrian", lastName: "Enriquez", office: "Tucson", position: "Laborer" },
  { id: "emp162", number: "1704", firstName: "Manuel", lastName: "Enriquez", office: "Tucson", position: "Foreman" },
  { id: "emp163", number: "2850", firstName: "Nathan", lastName: "Enriquez", office: "Sierra Vista", position: "Grade Checker" },
  { id: "emp164", number: "1983", firstName: "Teofilo", lastName: "Escarcega", office: "Tucson", position: "Laborer" },
  { id: "emp165", number: "3003", firstName: "Giovoni", lastName: "Escobedo", office: "Tucson", position: "Operator" },
  { id: "emp166", number: "2826", firstName: "Jesus", lastName: "Espinoza", office: "Tucson", position: "Pipelayer" },
  { id: "emp167", number: "2923", firstName: "Roberto", lastName: "Estrada", office: "Sierra Vista", position: "Laborer" },
  { id: "emp168", number: "3038", firstName: "Daniel", lastName: "Estrada Ramos", office: "Tucson", position: "Laborer" },
  { id: "emp169", number: "2654", firstName: "Santiago", lastName: "Estrella", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp170", number: "2927", firstName: "Aaron", lastName: "Evans", office: "Sierra Vista", position: "Laborer" },
  { id: "emp171", number: "1956", firstName: "Andrew", lastName: "Evans", office: "Sierra Vista", position: "Operator" },
  { id: "emp172", number: "2609", firstName: "Austin", lastName: "Evans", office: "Tucson", position: "Mechanic" },
  { id: "emp173", number: "53285", firstName: "Roberto", lastName: "Evans", office: "Sierra Vista", position: "" },
  { id: "emp174", number: "64391", firstName: "David", lastName: "Farnsworth", office: "Tucson", position: "Foreman" },
  { id: "emp175", number: "3132", firstName: "William", lastName: "Farrell Anderson", office: "Sierra Vista", position: "" },
  { id: "emp176", number: "45164", firstName: "Mark", lastName: "Faruolo", office: "Tucson", position: "" },
  { id: "emp177", number: "2623", firstName: "Francisco", lastName: "Figueroa", office: "Tucson", position: "Carpenter" },
  { id: "emp178", number: "3102", firstName: "Carissa", lastName: "Fincher", office: "Sierra Vista", position: "Field Engineer" },
  { id: "emp179", number: "1661", firstName: "Francisco", lastName: "Flores", office: "Tucson", position: "Operator" },
  { id: "emp180", number: "1008", firstName: "Daniel", lastName: "Fluno", office: "Tucson", position: "Pipelayer" },
  { id: "emp181", number: "1398", firstName: "Matthew", lastName: "Ford", office: "Sierra Vista", position: "Leadman" },
  { id: "emp182", number: "3009", firstName: "Mason", lastName: "Fosdick", office: "Tucson", position: "Engineering Intern" },
  { id: "emp183", number: "1240", firstName: "Jonathan", lastName: "Fraley Jr", office: "Tucson", position: "Operator" },
  { id: "emp184", number: "2600", firstName: "Jason", lastName: "Francis", office: "Tucson", position: "Leadman" },
  { id: "emp185", number: "1752", firstName: "Jose", lastName: "Franco", office: "Tucson", position: "Foreman" },
  { id: "emp186", number: "1945", firstName: "Jose", lastName: "Franco", office: "Tucson", position: "Pipelayer" },
  { id: "emp187", number: "1695", firstName: "Valeria", lastName: "Fuerte", office: "Tucson", position: "Project Engineer" },
  { id: "emp188", number: "2970", firstName: "Kasey", lastName: "Gaddy", office: "Tucson", position: "Director Of Communications" },
  { id: "emp189", number: "2771", firstName: "Refugio", lastName: "Galicia", office: "Tucson", position: "Pipelayer" },
  { id: "emp190", number: "2349", firstName: "Alberto", lastName: "Galindo", office: "Sierra Vista", position: "Laborer" },
  { id: "emp191", number: "97957", firstName: "Luis", lastName: "Galindo", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp192", number: "1159", firstName: "Elizabeth", lastName: "Gallego", office: "Tucson", position: "" },
  { id: "emp193", number: "72466", firstName: "Sergio", lastName: "Gallego", office: "Tucson", position: "" },
  { id: "emp194", number: "2995", firstName: "Anthony", lastName: "Galvan", office: "Tucson", position: "Laborer" },
  { id: "emp195", number: "2508", firstName: "Joaquin", lastName: "Gamez Arvayo", office: "Tucson", position: "Pipelayer" },
  { id: "emp196", number: "3048", firstName: "David", lastName: "Gandarela", office: "Tucson", position: "Project Engineer" },
  { id: "emp197", number: "3023", firstName: "Jose", lastName: "Garcia", office: "Tucson", position: "Engineering Intern" },
  { id: "emp198", number: "2251", firstName: "Joseph", lastName: "Garcia", office: "Tucson", position: "Operator" },
  { id: "emp199", number: "2820", firstName: "Juan", lastName: "Garcia", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp200", number: "2955", firstName: "Luis", lastName: "Garcia", office: "Sierra Vista", position: "Operator" },
  { id: "emp201", number: "2622", firstName: "Luis", lastName: "Garcia", office: "Tucson", position: "" },
  { id: "emp202", number: "1263", firstName: "Michael", lastName: "Garcia", office: "Tucson", position: "Superintendent" },
  { id: "emp203", number: "2949", firstName: "Rafael", lastName: "Garcia Samaniego", office: "Sierra Vista", position: "Laborer" },
  { id: "emp204", number: "2755", firstName: "Bardo", lastName: "Gastelum Hernandez", office: "Sierra Vista", position: "Laborer" },
  { id: "emp205", number: "2667", firstName: "Joseph", lastName: "Gauthier", office: "Tucson", position: "" },
  { id: "emp206", number: "5744", firstName: "Duane", lastName: "Giltner", office: "Tucson", position: "Operator" },
  { id: "emp207", number: "1637", firstName: "Nathan", lastName: "Goddard", office: "Tucson", position: "Shop Foreman" },
  { id: "emp208", number: "3039", firstName: "Manuel", lastName: "Godinez", office: "Tucson", position: "Pipelayer" },
  { id: "emp209", number: "1186", firstName: "Richard", lastName: "Godoy", office: "Tucson", position: "Laborer" },
  { id: "emp210", number: "2389", firstName: "Angel", lastName: "Gomez", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp211", number: "2963", firstName: "Carlos", lastName: "Gomez", office: "Tucson", position: "Operator" },
  { id: "emp212", number: "2611", firstName: "Sergio", lastName: "Gomez Cuevas", office: "Tucson", position: "Laborer" },
  { id: "emp213", number: "2549", firstName: "David", lastName: "Gonzales", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp214", number: "1574", firstName: "Richard", lastName: "Gonzales", office: "Tucson", position: "Operator" },
  { id: "emp215", number: "2008", firstName: "Alfred", lastName: "Gonzalez", office: "Tucson", position: "Foreman" },
  { id: "emp216", number: "3018", firstName: "Ali", lastName: "Gonzalez", office: "Tucson", position: "Laborer" },
  { id: "emp217", number: "2608", firstName: "Andres", lastName: "Gonzalez", office: "Tucson", position: "Operator" },
  { id: "emp218", number: "2982", firstName: "Jessica", lastName: "Gonzalez", office: "Tucson", position: "Operator" },
  { id: "emp219", number: "2928", firstName: "Jeffrey", lastName: "Goodman", office: "Sierra Vista", position: "Operator" },
  { id: "emp220", number: "3120", firstName: "Steven", lastName: "Goodman", office: "Sierra Vista", position: "Superintendent" },
  { id: "emp221", number: "2962", firstName: "Chazz", lastName: "Goodwin", office: "Tucson", position: "Operator" },
  { id: "emp222", number: "1563", firstName: "Raul", lastName: "Gracia", office: "Sierra Vista", position: "Laborer" },
  { id: "emp223", number: "2869", firstName: "Michael", lastName: "Grant", office: "Tucson", position: "Laborer" },
  { id: "emp224", number: "3106", firstName: "James", lastName: "Gregory", office: "Sierra Vista", position: "Laborer" },
  { id: "emp225", number: "2590", firstName: "Francisco", lastName: "Grijalva", office: "Sierra Vista", position: "Carpenter" },
  { id: "emp226", number: "2997", firstName: "Pedro", lastName: "Grijalva", office: "Tucson", position: "Operator" },
  { id: "emp227", number: "3135", firstName: "Hector", lastName: "Grijalva Hernandez", office: "Sierra Vista", position: "Laborer" },
  { id: "emp228", number: "1990", firstName: "James", lastName: "Grimes", office: "Tucson", position: "Field Engineer" },
  { id: "emp229", number: "1861", firstName: "Keith", lastName: "Gryczkowski", office: "Tucson", position: "Operator" },
  { id: "emp230", number: "2647", firstName: "Guillermo", lastName: "Guerrero", office: "Tucson", position: "Operator" },
  { id: "emp231", number: "2800", firstName: "Jose", lastName: "Guerrero", office: "Tucson", position: "Operator" },
  { id: "emp232", number: "2707", firstName: "Miguel", lastName: "Gutierrez Ortega", office: "Sierra Vista", position: "Pipelayer" },
  { id: "emp233", number: "1528", firstName: "Michael", lastName: "Harrison", office: "Tucson", position: "Mechanic" },
  { id: "emp234", number: "2153", firstName: "Dezmond", lastName: "Henry", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp235", number: "3006", firstName: "Dante", lastName: "Herbert", office: "Tucson", position: "Pipelayer" },
  { id: "emp236", number: "2718", firstName: "Alex", lastName: "Hernandez", office: "Sierra Vista", position: "Operator" },
  { id: "emp237", number: "2332", firstName: "Avidan", lastName: "Hernandez", office: "Tucson", position: "Operator" },
  { id: "emp238", number: "1256", firstName: "Cesar", lastName: "Hernandez", office: "Tucson", position: "Laborer" },
  { id: "emp239", number: "1080", firstName: "Edward", lastName: "Hernandez", office: "Tucson", position: "Operator" },
  { id: "emp240", number: "2563", firstName: "Fabian", lastName: "Hernandez", office: "Tucson", position: "Laborer" },
  { id: "emp241", number: "2910", firstName: "Jason", lastName: "Hernandez", office: "Tucson", position: "Foreman" },
  { id: "emp242", number: "2682", firstName: "Johnny", lastName: "Hernandez", office: "Tucson", position: "Laborer" },
  { id: "emp243", number: "2469", firstName: "Luis", lastName: "Hernandez", office: "Sierra Vista", position: "Operator" },
  { id: "emp244", number: "1673", firstName: "Ruben", lastName: "Hernandez", office: "Tucson", position: "Leadman" },
  { id: "emp245", number: "2436", firstName: "Aldo", lastName: "Herrera", office: "Tucson", position: "" },
  { id: "emp246", number: "2797", firstName: "Jasiel", lastName: "Higuera Velasquez", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp247", number: "2991", firstName: "Eric", lastName: "Hinton", office: "Tucson", position: "Field Engineer" },
  { id: "emp248", number: "2486", firstName: "Joshua", lastName: "Hinyup", office: "Tucson", position: "" },
  { id: "emp249", number: "2994", firstName: "Julian", lastName: "Holguin", office: "Tucson", position: "Laborer" },
  { id: "emp250", number: "78859", firstName: "Luis", lastName: "Holguin", office: "Tucson", position: "Foreman" },
  { id: "emp251", number: "3055", firstName: "Mythias", lastName: "Holguin", office: "Tucson", position: "Engineering Intern" },
  { id: "emp252", number: "2822", firstName: "Karl", lastName: "Holm", office: "Tucson", position: "Laborer" },
  { id: "emp253", number: "2513", firstName: "Chad", lastName: "Housler", office: "Tucson", position: "Chief Estimator" },
  { id: "emp254", number: "2305", firstName: "Justin", lastName: "Housler", office: "Tucson", position: "" },
  { id: "emp255", number: "2648", firstName: "Adrian", lastName: "Hoyos Monge", office: "Sierra Vista", position: "Laborer" },
  { id: "emp256", number: "1745", firstName: "Alejandro", lastName: "Hudson Martinez", office: "Tucson", position: "Cdl Driver" },
  { id: "emp257", number: "2522", firstName: "Ernesto", lastName: "Huerta", office: "Sierra Vista", position: "Operator" },
  { id: "emp258", number: "1157", firstName: "Daniel", lastName: "Huff", office: "Tucson", position: "Operator" },
  { id: "emp259", number: "1694", firstName: "Jose", lastName: "Ibarra", office: "Tucson", position: "Cdl Driver" },
  { id: "emp260", number: "2738", firstName: "Raul", lastName: "Inclan", office: "Tucson", position: "Operator" },
  { id: "emp261", number: "2352", firstName: "Jagger", lastName: "Jacobi", office: "Tucson", position: "Field Engineer" },
  { id: "emp262", number: "2898", firstName: "Neimiah", lastName: "Joel", office: "Tucson", position: "Laborer" },
  { id: "emp263", number: "3030", firstName: "Al", lastName: "Jones", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp264", number: "2129", firstName: "Brett", lastName: "Jones", office: "Tucson", position: "Field Safety Manager" },
  { id: "emp265", number: "2203", firstName: "Derek", lastName: "Jones", office: "Tucson", position: "" },
  { id: "emp266", number: "608", firstName: "Dannie", lastName: "Jordan", office: "Tucson", position: "Superintendent" },
  { id: "emp267", number: "2900", firstName: "Wyatt", lastName: "Jozwiak", office: "Tucson", position: "Laborer" },
  { id: "emp268", number: "2833", firstName: "Diane", lastName: "Kawada", office: "Sierra Vista", position: "Project Administrator" },
  { id: "emp269", number: "2565", firstName: "Johnathon", lastName: "Kelly", office: "Tucson", position: "" },
  { id: "emp270", number: "2712", firstName: "Caleb", lastName: "King", office: "Sierra Vista", position: "Operator" },
  { id: "emp271", number: "89605", firstName: "Clifton", lastName: "Klamath", office: "Sierra Vista", position: "Foreman" },
  { id: "emp272", number: "2542", firstName: "Hanna", lastName: "Krause", office: "Tucson", position: "Project Administrator" },
  { id: "emp273", number: "2819", firstName: "Joshua", lastName: "Krause", office: "Tucson", position: "Field Engineer" },
  { id: "emp274", number: "1370", firstName: "Christina", lastName: "Kuzyk", office: "Tucson", position: "Safety Director" },
  { id: "emp275", number: "2922", firstName: "Ian", lastName: "Lamma", office: "Sierra Vista", position: "Operator" },
  { id: "emp276", number: "2849", firstName: "Ryan", lastName: "Lamma", office: "Sierra Vista", position: "Laborer" },
  { id: "emp277", number: "2684", firstName: "Daniel", lastName: "Lane", office: "Tucson", position: "Superintendent" },
  { id: "emp278", number: "1502", firstName: "Marcel", lastName: "Lara", office: "Tucson", position: "Leadman" },
  { id: "emp279", number: "28610", firstName: "Marcelino", lastName: "Lara", office: "Tucson", position: "Operator" },
  { id: "emp280", number: "2529", firstName: "Ramon", lastName: "Lara", office: "Tucson", position: "Operator" },
  { id: "emp281", number: "2773", firstName: "Roman", lastName: "Larribas Perez", office: "Tucson", position: "Laborer" },
  { id: "emp282", number: "43152", firstName: "Jaimie", lastName: "Lavallee", office: "Sierra Vista", position: "Shop Foreman" },
  { id: "emp283", number: "2677", firstName: "Edgar", lastName: "Leal", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp284", number: "3111", firstName: "Jorge", lastName: "Lechuga", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp285", number: "2679", firstName: "Johnnie", lastName: "Lee", office: "Tucson", position: "Operator" },
  { id: "emp286", number: "2998", firstName: "Roman", lastName: "Lemburg", office: "Tucson", position: "Laborer" },
  { id: "emp287", number: "2336", firstName: "Collin", lastName: "Lewis", office: "Tucson", position: "Project Engineer" },
  { id: "emp288", number: "2945", firstName: "Nicholas", lastName: "Lewis", office: "Sierra Vista", position: "Laborer" },
  { id: "emp289", number: "2915", firstName: "Jose Guadalupe", lastName: "Leyva Nido", office: "Tucson", position: "Laborer" },
  { id: "emp290", number: "2671", firstName: "Martin", lastName: "Leyva Pacheco", office: "Tucson", position: "Operator" },
  { id: "emp291", number: "2916", firstName: "Scott", lastName: "Looney", office: "Tucson", position: "Operator" },
  { id: "emp292", number: "2953", firstName: "Camilo", lastName: "Lopez", office: "Sierra Vista", position: "Operator" },
  { id: "emp293", number: "81910", firstName: "Francisco", lastName: "Lopez", office: "Tucson", position: "Leadman" },
  { id: "emp294", number: "2789", firstName: "Gabriel", lastName: "Lopez", office: "Tucson", position: "" },
  { id: "emp295", number: "2059", firstName: "Eduardo", lastName: "Lopez Castro", office: "Sierra Vista", position: "Carpenter" },
  { id: "emp296", number: "3037", firstName: "Sergi", lastName: "Lopez Jr", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp297", number: "2592", firstName: "Cesar", lastName: "Lopez-Guerrero", office: "Sierra Vista", position: "Laborer" },
  { id: "emp298", number: "1521", firstName: "Dustin", lastName: "Lowe", office: "Sierra Vista", position: "Superintendent" },
  { id: "emp299", number: "2453", firstName: "Erik", lastName: "Loya", office: "Tucson", position: "Engineering Intern" },
  { id: "emp300", number: "1632", firstName: "Gabriel", lastName: "Luera", office: "Tucson", position: "Pipelayer" },
  { id: "emp301", number: "3000", firstName: "Miguel", lastName: "Lugo", office: "Tucson", position: "Field Engineer" },
  { id: "emp302", number: "3043", firstName: "Dagoberto", lastName: "Lujan", office: "Tucson", position: "" },
  { id: "emp303", number: "2328", firstName: "Doroteo", lastName: "Luna Valenzuela", office: "Tucson", position: "" },
  { id: "emp304", number: "2700", firstName: "John", lastName: "Macdougall", office: "Sierra Vista", position: "Project Manager_Estimator" },
  { id: "emp305", number: "2870", firstName: "Gabriel", lastName: "Machado", office: "Tucson", position: "" },
  { id: "emp306", number: "2658", firstName: "Isaac", lastName: "Madrid", office: "Tucson", position: "Superintendent" },
  { id: "emp307", number: "2741", firstName: "Edwin", lastName: "Madrigal Olivas", office: "Tucson", position: "Laborer" },
  { id: "emp308", number: "2704", firstName: "Hessgar", lastName: "Majalca", office: "Sierra Vista", position: "Laborer" },
  { id: "emp309", number: "3131", firstName: "Adam", lastName: "Malley", office: "Sierra Vista", position: "Laborer" },
  { id: "emp310", number: "2626", firstName: "Broden", lastName: "Marshall", office: "Tucson", position: "" },
  { id: "emp311", number: "2298", firstName: "Anthony", lastName: "Martinez", office: "Tucson", position: "" },
  { id: "emp312", number: "2804", firstName: "Arnold", lastName: "Martinez", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp313", number: "12756", firstName: "Ciriaco", lastName: "Martinez", office: "Sierra Vista", position: "Foreman" },
  { id: "emp314", number: "3004", firstName: "Gildardo", lastName: "Martinez", office: "Tucson", position: "Pipelayer" },
  { id: "emp315", number: "2807", firstName: "Isaac", lastName: "Martinez", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp316", number: "2909", firstName: "Jared", lastName: "Martinez", office: "Tucson", position: "Laborer" },
  { id: "emp317", number: "2517", firstName: "Marcos", lastName: "Martinez", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp318", number: "2830", firstName: "Nicholas", lastName: "Martinez", office: "Tucson", position: "Laborer" },
  { id: "emp319", number: "1489", firstName: "Ruben", lastName: "Martinez", office: "Tucson", position: "Carpenter" },
  { id: "emp320", number: "3019", firstName: "Gabriel", lastName: "Martinez Pico", office: "Tucson", position: "Laborer" },
  { id: "emp321", number: "2988", firstName: "Prabhjeet", lastName: "Matharoo", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp322", number: "2643", firstName: "Aaron", lastName: "Mathis", office: "Tucson", position: "Laborer" },
  { id: "emp323", number: "1103", firstName: "Rita", lastName: "Maxwell", office: "Tucson", position: "" },
  { id: "emp324", number: "2435", firstName: "Miguel", lastName: "Mazon", office: "Tucson", position: "Operator" },
  { id: "emp325", number: "75376", firstName: "Ricky", lastName: "Mcmahon", office: "Sierra Vista", position: "Superintendent" },
  { id: "emp326", number: "3020", firstName: "Aaron", lastName: "Mcpike", office: "Tucson", position: "Laborer" },
  { id: "emp327", number: "1185", firstName: "Alonzo", lastName: "Medina", office: "Tucson", position: "Pipelayer" },
  { id: "emp328", number: "2446", firstName: "Luis", lastName: "Medina", office: "Sierra Vista", position: "Operator" },
  { id: "emp329", number: "1448", firstName: "Steve", lastName: "Meeks", office: "Tucson", position: "Superintendent" },
  { id: "emp330", number: "2908", firstName: "Max", lastName: "Mejia", office: "Tucson", position: "Superintendent" },
  { id: "emp331", number: "2853", firstName: "Joshua", lastName: "Mena", office: "Tucson", position: "Laborer" },
  { id: "emp332", number: "1670", firstName: "Francisco", lastName: "Mendez", office: "Sierra Vista", position: "Operator" },
  { id: "emp333", number: "2617", firstName: "Joshua", lastName: "Mendez", office: "Tucson", position: "Laborer" },
  { id: "emp334", number: "1276", firstName: "Christopher", lastName: "Mendivil", office: "Sierra Vista", position: "Leadman" },
  { id: "emp335", number: "2394", firstName: "Francisco", lastName: "Mendoza", office: "Tucson", position: "Operator" },
  { id: "emp336", number: "3054", firstName: "Martha", lastName: "Mendoza", office: "Tucson", position: "Field Technician" },
  { id: "emp337", number: "1730", firstName: "Victor", lastName: "Mendoza", office: "Sierra Vista", position: "Pipelayer" },
  { id: "emp338", number: "1590", firstName: "Emmett", lastName: "Merchant", office: "Tucson", position: "General Superintendent" },
  { id: "emp339", number: "3040", firstName: "Eddie", lastName: "Miguel", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp340", number: "1557", firstName: "Shaun", lastName: "Miller", office: "Tucson", position: "" },
  { id: "emp341", number: "2173", firstName: "Devin", lastName: "Million", office: "Tucson", position: "Safety Assistant" },
  { id: "emp342", number: "77705", firstName: "George", lastName: "Miranda", office: "Tucson", position: "Superintendent" },
  { id: "emp343", number: "1899", firstName: "Daniel", lastName: "Mohineda", office: "Tucson", position: "Leadman" },
  { id: "emp344", number: "1820", firstName: "Pablo", lastName: "Mojarro", office: "Tucson", position: "Pipelayer" },
  { id: "emp345", number: "2764", firstName: "Bryan", lastName: "Molina Fimbres", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp346", number: "2605", firstName: "Alejandro", lastName: "Monge", office: "Tucson", position: "Leadman" },
  { id: "emp347", number: "2297", firstName: "Victor", lastName: "Monjaras", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp348", number: "3137", firstName: "Alexandro", lastName: "Montano Aguilar", office: "Sierra Vista", position: "Laborer" },
  { id: "emp349", number: "3139", firstName: "Francisco", lastName: "Montano Aguilar", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp350", number: "3138", firstName: "Jesus", lastName: "Montano-Aguilar", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp351", number: "2767", firstName: "Christian", lastName: "Montez", office: "Sierra Vista", position: "Leadman" },
  { id: "emp352", number: "2398", firstName: "Michael", lastName: "Montgomery Jr", office: "Tucson", position: "Operator" },
  { id: "emp353", number: "2846", firstName: "Gabriel", lastName: "Montiel", office: "Sierra Vista", position: "Laborer" },
  { id: "emp354", number: "2669", firstName: "Gerardo", lastName: "Montiel Sanchez", office: "Tucson", position: "Operator" },
  { id: "emp355", number: "2597", firstName: "Augustine", lastName: "Montoya", office: "Sierra Vista", position: "Laborer" },
  { id: "emp356", number: "2492", firstName: "Mark", lastName: "Montoya", office: "Tucson", position: "Laborer" },
  { id: "emp357", number: "3051", firstName: "Walter", lastName: "Montoya", office: "Tucson", position: "Laborer" },
  { id: "emp358", number: "2835", firstName: "Angel", lastName: "Morales", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp359", number: "2692", firstName: "Lilian", lastName: "Morales", office: "Tucson", position: "" },
  { id: "emp360", number: "2987", firstName: "Peter", lastName: "Morales", office: "Tucson", position: "Field Engineer" },
  { id: "emp361", number: "36706", firstName: "Ramon", lastName: "Morales", office: "Sierra Vista", position: "Superintendent" },
  { id: "emp362", number: "2644", firstName: "Carlos", lastName: "Morales Barcelo", office: "Tucson", position: "Operator" },
  { id: "emp363", number: "2662", firstName: "Francisco", lastName: "Morales Gutierrez", office: "Tucson", position: "Pipelayer" },
  { id: "emp364", number: "2989", firstName: "Alexia", lastName: "Morales Ortiz", office: "Tucson", position: "Field Engineer" },
  { id: "emp365", number: "2951", firstName: "Adam", lastName: "Moreno", office: "Sierra Vista", position: "Laborer" },
  { id: "emp366", number: "2221", firstName: "Antonio", lastName: "Moreno", office: "Tucson", position: "Project Engineer" },
  { id: "emp367", number: "1302", firstName: "Javier", lastName: "Moreno", office: "Sierra Vista", position: "Superintendent" },
  { id: "emp368", number: "2578", firstName: "Victor", lastName: "Moreno", office: "Sierra Vista", position: "Pipelayer" },
  { id: "emp369", number: "1992", firstName: "Luis", lastName: "Morgan Sanchez", office: "Tucson", position: "Pipelayer" },
  { id: "emp370", number: "2326", firstName: "Jason", lastName: "Mullins", office: "Sierra Vista", position: "Laborer" },
  { id: "emp371", number: "2933", firstName: "Brent", lastName: "Mulvaney", office: "Sierra Vista", position: "Operator" },
  { id: "emp372", number: "2618", firstName: "Levi", lastName: "Mulverhill", office: "Tucson", position: "Field Engineer" },
  { id: "emp373", number: "1202", firstName: "John", lastName: "Mulverhill Jr", office: "Tucson", position: "Senior Superintendent" },
  { id: "emp374", number: "2803", firstName: "Jose", lastName: "Munguia Quijada", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp375", number: "2958", firstName: "Alexis", lastName: "Munoz-Almanza", office: "Sierra Vista", position: "Laborer" },
  { id: "emp376", number: "2639", firstName: "Jesus", lastName: "Murguia", office: "Tucson", position: "Operator" },
  { id: "emp377", number: "1969", firstName: "Filiberto", lastName: "Murillo", office: "Tucson", position: "Leadman" },
  { id: "emp378", number: "86404", firstName: "Nathan", lastName: "Murken", office: "Sierra Vista", position: "" },
  { id: "emp379", number: "2537", firstName: "Teryl", lastName: "Murray", office: "Tucson", position: "Division Manager" },
  { id: "emp380", number: "2956", firstName: "Michael", lastName: "Negrette", office: "Sierra Vista", position: "Laborer" },
  { id: "emp381", number: "1156", firstName: "Matthew", lastName: "Nehrmeyer", office: "Tucson", position: "Division Manager" },
  { id: "emp382", number: "2270", firstName: "Sheana", lastName: "Nidetz", office: "Tucson", position: "" },
  { id: "emp383", number: "2539", firstName: "Johanna", lastName: "Noon", office: "Tucson", position: "" },
  { id: "emp384", number: "1511", firstName: "Edward", lastName: "Norris", office: "Tucson", position: "Superintendent" },
  { id: "emp385", number: "57619", firstName: "George", lastName: "Norwood", office: "Sierra Vista", position: "Foreman" },
  { id: "emp386", number: "2458", firstName: "Juan", lastName: "Nunez Garcia", office: "Tucson", position: "Pipelayer" },
  { id: "emp387", number: "2350", firstName: "Luis", lastName: "Nunez Montes", office: "Sierra Vista", position: "Laborer" },
  { id: "emp388", number: "2917", firstName: "Diego", lastName: "Ochoa", office: "Tucson", position: "Operator" },
  { id: "emp389", number: "2427", firstName: "Matthew", lastName: "Ogle", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp390", number: "2591", firstName: "Raudel", lastName: "Olague Vazquez", office: "Sierra Vista", position: "Laborer" },
  { id: "emp391", number: "2978", firstName: "Edward", lastName: "Olejnik", office: "Tucson", position: "Pipelayer" },
  { id: "emp392", number: "45200", firstName: "Jeff", lastName: "Olejnik", office: "Tucson", position: "" },
  { id: "emp393", number: "65681", firstName: "Nicholas", lastName: "Olejnik", office: "Tucson", position: "Senior Superintendent" },
  { id: "emp394", number: "3002", firstName: "Rudy", lastName: "Olivas", office: "Tucson", position: "Operator" },
  { id: "emp395", number: "1155", firstName: "James", lastName: "Olson", office: "Tucson", position: "Division Manager" },
  { id: "emp396", number: "2628", firstName: "Armando", lastName: "Orduno", office: "Tucson", position: "Foreman" },
  { id: "emp397", number: "2649", firstName: "Larry", lastName: "Orduno Flores", office: "Tucson", position: "Operator" },
  { id: "emp398", number: "1144", firstName: "Santiago", lastName: "Orozco", office: "Tucson", position: "Cdl Driver" },
  { id: "emp399", number: "2772", firstName: "Jesse", lastName: "Ortega", office: "Tucson", position: "Pipelayer" },
  { id: "emp400", number: "2329", firstName: "Bernando", lastName: "Ortiz", office: "Tucson", position: "Laborer" },
  { id: "emp401", number: "28544", firstName: "Rafael", lastName: "Ortiz", office: "Sierra Vista", position: "Laborer" },
  { id: "emp402", number: "3041", firstName: "Javier", lastName: "Ortiz Jr", office: "Tucson", position: "Operator" },
  { id: "emp403", number: "2653", firstName: "Victor", lastName: "Ortiz Vazquez", office: "Tucson", position: "Foreman" },
  { id: "emp404", number: "1294", firstName: "Cheyenne", lastName: "Padilla", office: "Tucson", position: "Leadman" },
  { id: "emp405", number: "2689", firstName: "Isaiah", lastName: "Padilla", office: "Tucson", position: "Field Engineer" },
  { id: "emp406", number: "2621", firstName: "Robert", lastName: "Palma", office: "Tucson", position: "Laborer" },
  { id: "emp407", number: "8598", firstName: "Ross", lastName: "Parker", office: "Sierra Vista", position: "Superintendent" },
  { id: "emp408", number: "2943", firstName: "Leonel", lastName: "Parra", office: "Sierra Vista", position: "Laborer" },
  { id: "emp409", number: "1461", firstName: "Lawrence", lastName: "Paxton", office: "Tucson", position: "Laborer" },
  { id: "emp410", number: "3013", firstName: "Isberlayne", lastName: "Pena", office: "Tucson", position: "Laborer" },
  { id: "emp411", number: "12297", firstName: "Angel", lastName: "Perez", office: "Tucson", position: "Leadman" },
  { id: "emp412", number: "2996", firstName: "Rigo", lastName: "Perez", office: "Tucson", position: "Laborer" },
  { id: "emp413", number: "2239", firstName: "Robert", lastName: "Perkins", office: "Tucson", position: "" },
  { id: "emp414", number: "1911", firstName: "Ronny", lastName: "Perpuly", office: "Tucson", position: "Leadman" },
  { id: "emp415", number: "3109", firstName: "Brysonn", lastName: "Perry", office: "Sierra Vista", position: "Laborer" },
  { id: "emp416", number: "3119", firstName: "Johnny", lastName: "Peru", office: "Sierra Vista", position: "Laborer" },
  { id: "emp417", number: "2576", firstName: "Amanda", lastName: "Peters", office: "Sierra Vista", position: "Project Engineer" },
  { id: "emp418", number: "2657", firstName: "Nicholas", lastName: "Petersen", office: "Tucson", position: "Superintendent" },
  { id: "emp419", number: "2891", firstName: "Christopher", lastName: "Pino", office: "Tucson", position: "Operator" },
  { id: "emp420", number: "3101", firstName: "Michael", lastName: "Porras", office: "Sierra Vista", position: "Laborer" },
  { id: "emp421", number: "2742", firstName: "John", lastName: "Portillo", office: "Tucson", position: "Project Engineer" },
  { id: "emp422", number: "2391", firstName: "Manuel", lastName: "Portugal", office: "Tucson", position: "Project Engineer" },
  { id: "emp423", number: "2969", firstName: "Ernest", lastName: "Posada", office: "Tucson", position: "Operator" },
  { id: "emp424", number: "1491", firstName: "Kendra", lastName: "Puhala", office: "Tucson", position: "Field Engineer" },
  { id: "emp425", number: "1312", firstName: "Debra", lastName: "Pursley", office: "Tucson", position: "Operator" },
  { id: "emp426", number: "2574", firstName: "Toby", lastName: "Putman", office: "Tucson", position: "Carpenter" },
  { id: "emp427", number: "2877", firstName: "Danielle", lastName: "Quenzler", office: "Tucson", position: "" },
  { id: "emp428", number: "3115", firstName: "Alvaro", lastName: "Quihui Cohen", office: "Sierra Vista", position: "Laborer" },
  { id: "emp429", number: "2959", firstName: "Amado", lastName: "Quijada Jr", office: "Sierra Vista", position: "Laborer" },
  { id: "emp430", number: "2377", firstName: "Jesus", lastName: "Quiroz", office: "Sierra Vista", position: "Project Engineer" },
  { id: "emp431", number: "2693", firstName: "Lisa", lastName: "Rahe", office: "Tucson", position: "" },
  { id: "emp432", number: "2729", firstName: "Diego", lastName: "Ramirez", office: "Tucson", position: "" },
  { id: "emp433", number: "1501", firstName: "Jody", lastName: "Ramirez", office: "Tucson", position: "Laborer" },
  { id: "emp434", number: "2874", firstName: "Joseph", lastName: "Ramirez", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp435", number: "2896", firstName: "Lucy", lastName: "Ramirez", office: "Tucson", position: "" },
  { id: "emp436", number: "3112", firstName: "Rey", lastName: "Ramirez Lopez", office: "Sierra Vista", position: "Operator" },
  { id: "emp437", number: "2687", firstName: "Alex", lastName: "Ramos", office: "Tucson", position: "Engineering Intern" },
  { id: "emp438", number: "2625", firstName: "Susano", lastName: "Ramos", office: "Tucson", position: "Operator" },
  { id: "emp439", number: "1982", firstName: "Leonard", lastName: "Ramsey", office: "Tucson", position: "Foreman" },
  { id: "emp440", number: "2585", firstName: "Michael", lastName: "Rangel", office: "Sierra Vista", position: "Laborer" },
  { id: "emp441", number: "2753", firstName: "Raziel", lastName: "Rascon-Mendez", office: "Sierra Vista", position: "Operator" },
  { id: "emp442", number: "2086", firstName: "Charles", lastName: "Ratliff", office: "Tucson", position: "Foreman" },
  { id: "emp443", number: "2187", firstName: "Matthew", lastName: "Ratliff", office: "Tucson", position: "Operator" },
  { id: "emp444", number: "3130", firstName: "Dylan", lastName: "Ray", office: "Sierra Vista", position: "Laborer" },
  { id: "emp445", number: "96811", firstName: "Samuel", lastName: "Recinos", office: "Tucson", position: "Foreman" },
  { id: "emp446", number: "2890", firstName: "David", lastName: "Reed", office: "Tucson", position: "Laborer" },
  { id: "emp447", number: "2944", firstName: "Tyler", lastName: "Reed", office: "Sierra Vista", position: "Operator" },
  { id: "emp448", number: "2986", firstName: "Patrick", lastName: "Reilly", office: "Tucson", position: "Cdl Driver" },
  { id: "emp449", number: "2018", firstName: "Raymond", lastName: "Reyes", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp450", number: "1389", firstName: "Stephen", lastName: "Rice", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp451", number: "2686", firstName: "Devin", lastName: "Richardson", office: "Tucson", position: "Mechanic" },
  { id: "emp452", number: "2154", firstName: "Zerik", lastName: "Richardson", office: "Tucson", position: "Mechanic" },
  { id: "emp453", number: "2971", firstName: "Gabriel", lastName: "Ridenour", office: "Tucson", position: "Laborer" },
  { id: "emp454", number: "2905", firstName: "Jason", lastName: "Ridgway", office: "Tucson", position: "Laborer" },
  { id: "emp455", number: "1392", firstName: "Richard", lastName: "Riggleman", office: "Tucson", position: "Operator" },
  { id: "emp456", number: "2255", firstName: "Angel", lastName: "Rios", office: "Tucson", position: "Operator" },
  { id: "emp457", number: "2710", firstName: "Jose", lastName: "Rios Cardenas", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp458", number: "2630", firstName: "Jose", lastName: "Rivas Miranda", office: "Tucson", position: "Laborer" },
  { id: "emp459", number: "3121", firstName: "Roberto", lastName: "Rivera", office: "Sierra Vista", position: "Pipelayer" },
  { id: "emp460", number: "2145", firstName: "James", lastName: "Roberts", office: "Tucson", position: "Operator" },
  { id: "emp461", number: "2212", firstName: "Eric", lastName: "Robles", office: "Tucson", position: "Operator" },
  { id: "emp462", number: "3117", firstName: "Emilio", lastName: "Rodriguez", office: "Sierra Vista", position: "Operator" },
  { id: "emp463", number: "3025", firstName: "Gerardo", lastName: "Rodriguez", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp464", number: "2227", firstName: "Hector", lastName: "Rodriguez", office: "Tucson", position: "Operator" },
  { id: "emp465", number: "2113", firstName: "Jose", lastName: "Rodriguez", office: "Tucson", position: "Foreman" },
  { id: "emp466", number: "2296", firstName: "Sergio", lastName: "Rodriguez Cabrales", office: "Tucson", position: "" },
  { id: "emp467", number: "2463", firstName: "Luis", lastName: "Rodriguez Puerto", office: "Sierra Vista", position: "Laborer" },
  { id: "emp468", number: "2783", firstName: "Braulio", lastName: "Rodriguez Roman", office: "Tucson", position: "Pipelayer" },
  { id: "emp469", number: "1244", firstName: "Jesus", lastName: "Rodriguez Sanchez", office: "Tucson", position: "Foreman" },
  { id: "emp470", number: "2836", firstName: "Dallen", lastName: "Rogers", office: "Sierra Vista", position: "Laborer" },
  { id: "emp471", number: "2257", firstName: "Ramon", lastName: "Rojo Alvarez", office: "Tucson", position: "Pipelayer" },
  { id: "emp472", number: "2793", firstName: "Alex", lastName: "Romero", office: "Tucson", position: "Operator" },
  { id: "emp473", number: "2292", firstName: "Daniel", lastName: "Romero", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp474", number: "2975", firstName: "Alonso", lastName: "Romo", office: "Tucson", position: "Laborer" },
  { id: "emp475", number: "2694", firstName: "Justin", lastName: "Rosales", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp476", number: "2246", firstName: "Brandon", lastName: "Rosson", office: "Tucson", position: "Operator" },
  { id: "emp477", number: "2218", firstName: "Clayton", lastName: "Rosson", office: "Tucson", position: "Operator" },
  { id: "emp478", number: "68149", firstName: "Cuauhtemoc", lastName: "Rubio", office: "Tucson", position: "Shop Laborer" },
  { id: "emp479", number: "3005", firstName: "Benjamin", lastName: "Ruiz", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp480", number: "2367", firstName: "Jose", lastName: "Ruiz", office: "Sierra Vista", position: "Laborer" },
  { id: "emp481", number: "3027", firstName: "Mark", lastName: "Ruiz", office: "Tucson", position: "Operator" },
  { id: "emp482", number: "3028", firstName: "Tanya", lastName: "Ruiz", office: "Tucson", position: "Operator" },
  { id: "emp483", number: "2365", firstName: "Thomas", lastName: "Ruiz", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp484", number: "2920", firstName: "Anthony", lastName: "Rye", office: "Tucson", position: "Laborer" },
  { id: "emp485", number: "2596", firstName: "Jesus", lastName: "Saavedra", office: "Sierra Vista", position: "Mechanic" },
  { id: "emp486", number: "2999", firstName: "Francisco", lastName: "Saenz", office: "Tucson", position: "Pipelayer" },
  { id: "emp487", number: "3122", firstName: "Tadeo", lastName: "Samaniego Olivares", office: "Sierra Vista", position: "Laborer" },
  { id: "emp488", number: "1488", firstName: "Efraim", lastName: "Sanchez", office: "Tucson", position: "Leadman" },
  { id: "emp489", number: "2467", firstName: "Jose", lastName: "Sanchez", office: "Sierra Vista", position: "Laborer" },
  { id: "emp490", number: "2892", firstName: "Jose Eulogio", lastName: "Sanchez", office: "Tucson", position: "Laborer" },
  { id: "emp491", number: "3007", firstName: "Santos", lastName: "Sanchez", office: "Tucson", position: "Mechanic" },
  { id: "emp492", number: "1433", firstName: "Sergio", lastName: "Sanchez", office: "Tucson", position: "Foreman" },
  { id: "emp493", number: "1755", firstName: "Crysthian", lastName: "Sanchez Rivera", office: "Tucson", position: "Pipelayer" },
  { id: "emp494", number: "62773", firstName: "David", lastName: "Sandusky", office: "Tucson", position: "Mechanic" },
  { id: "emp495", number: "2906", firstName: "Antonio", lastName: "Santa Cruz", office: "Tucson", position: "Pipelayer" },
  { id: "emp496", number: "13232", firstName: "Armando", lastName: "Santa Maria", office: "Tucson", position: "Operator" },
  { id: "emp497", number: "1058", firstName: "Luis", lastName: "Santa Maria", office: "Tucson", position: "Operator" },
  { id: "emp498", number: "33476", firstName: "Roderico", lastName: "Santa Maria", office: "Tucson", position: "Leadman" },
  { id: "emp499", number: "92072", firstName: "Federico", lastName: "Santamaria", office: "Tucson", position: "Foreman" },
  { id: "emp500", number: "47903", firstName: "Jesus", lastName: "Santamaria", office: "Tucson", position: "Foreman" },
  { id: "emp501", number: "13024", firstName: "Jesus", lastName: "Santamaria", office: "Tucson", position: "Foreman" },
  { id: "emp502", number: "2942", firstName: "Jakob", lastName: "Schott", office: "Sierra Vista", position: "Laborer" },
  { id: "emp503", number: "1510", firstName: "Clinton", lastName: "Schweigert", office: "Tucson", position: "Equipment Manager" },
  { id: "emp504", number: "2089", firstName: "Jacob", lastName: "Schweigert", office: "Tucson", position: "Mechanic" },
  { id: "emp505", number: "2374", firstName: "Anthony", lastName: "Seage", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp506", number: "2713", firstName: "Nicholas", lastName: "Seifert", office: "Sierra Vista", position: "Laborer" },
  { id: "emp507", number: "2790", firstName: "Ampelio", lastName: "Sepulveda", office: "Tucson", position: "Concrete Finisher" },
  { id: "emp508", number: "2445", firstName: "Manuel", lastName: "Serrano", office: "Sierra Vista", position: "Leadman" },
  { id: "emp509", number: "2861", firstName: "Mark", lastName: "Shearer", office: "Tucson", position: "Controller" },
  { id: "emp510", number: "2544", firstName: "Nathan", lastName: "Shechter", office: "Tucson", position: "Superintendent" },
  { id: "emp511", number: "1052", firstName: "Kyle", lastName: "Shepherd", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp512", number: "2841", firstName: "Isaac", lastName: "Sheppard", office: "Sierra Vista", position: "Field Engineer" },
  { id: "emp513", number: "2708", firstName: "Adrian", lastName: "Silva", office: "Sierra Vista", position: "" },
  { id: "emp514", number: "1544", firstName: "Jesus", lastName: "Simpson", office: "Tucson", position: "Pipelayer" },
  { id: "emp515", number: "2524", firstName: "Troy", lastName: "Simpson", office: "Sierra Vista", position: "Foreman" },
  { id: "emp516", number: "3105", firstName: "Tucker", lastName: "Smith", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp517", number: "2190", firstName: "Francisco", lastName: "Somoza Felix", office: "Sierra Vista", position: "Carpenter" },
  { id: "emp518", number: "2966", firstName: "Roberto", lastName: "Soqui", office: "Tucson", position: "Laborer" },
  { id: "emp519", number: "2761", firstName: "Francisco", lastName: "Sosa Cornejo", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp520", number: "2382", firstName: "Manuel", lastName: "Soto", office: "Tucson", position: "Laborer" },
  { id: "emp521", number: "2743", firstName: "Ruben", lastName: "Soto", office: "Tucson", position: "Field Engineer" },
  { id: "emp522", number: "1758", firstName: "Luis", lastName: "Soto Gonzalez", office: "Tucson", position: "Leadman" },
  { id: "emp523", number: "1970", firstName: "Sheri", lastName: "Stamps", office: "Tucson", position: "" },
  { id: "emp524", number: "2780", firstName: "Thomas", lastName: "Stephens", office: "Tucson", position: "Division Manager" },
  { id: "emp525", number: "2455", firstName: "Kevin", lastName: "Stewart", office: "Tucson", position: "Laborer" },
  { id: "emp526", number: "3053", firstName: "John", lastName: "Stiffey", office: "Tucson", position: "" },
  { id: "emp527", number: "2808", firstName: "Joshua", lastName: "Stone", office: "Tucson", position: "Operator" },
  { id: "emp528", number: "2385", firstName: "Carlos", lastName: "Suarez", office: "Tucson", position: "Controller" },
  { id: "emp529", number: "2343", firstName: "Jonathon", lastName: "Summers", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp530", number: "2967", firstName: "Micheal", lastName: "Summers", office: "Tucson", position: "Cdl Driver" },
  { id: "emp531", number: "2872", firstName: "Teresa", lastName: "Summers", office: "Tucson", position: "" },
  { id: "emp532", number: "1335", firstName: "Michael", lastName: "Tadeo", office: "Tucson", position: "Division Manager" },
  { id: "emp533", number: "2280", firstName: "Edward", lastName: "Talone", office: "Tucson", position: "Pipelayer" },
  { id: "emp534", number: "2462", firstName: "Elmer", lastName: "Tarazon", office: "Tucson", position: "" },
  { id: "emp535", number: "2399", firstName: "Martin", lastName: "Tarazon Moreno", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp536", number: "73383", firstName: "Samuel", lastName: "Tariku- Deriba", office: "Tucson", position: "Foreman" },
  { id: "emp537", number: "2728", firstName: "Stanley", lastName: "Tate", office: "Tucson", position: "Laborer" },
  { id: "emp538", number: "17815", firstName: "David", lastName: "Tautimer", office: "Sierra Vista", position: "Foreman" },
  { id: "emp539", number: "2023", firstName: "Steven", lastName: "Teichert", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp540", number: "2990", firstName: "Heriberto", lastName: "Teran", office: "Tucson", position: "Laborer" },
  { id: "emp541", number: "2080", firstName: "Larry", lastName: "Tinsley", office: "Sierra Vista", position: "Laborer" },
  { id: "emp542", number: "2828", firstName: "Ryan", lastName: "Tollison", office: "Tucson", position: "Pipelayer" },
  { id: "emp543", number: "2883", firstName: "Gustavo", lastName: "Torres", office: "Tucson", position: "Field Engineer" },
  { id: "emp544", number: "3133", firstName: "Victor", lastName: "Torres", office: "Sierra Vista", position: "" },
  { id: "emp545", number: "2581", firstName: "Ruben", lastName: "Torres Lagarda", office: "Sierra Vista", position: "Laborer" },
  { id: "emp546", number: "2523", firstName: "Francisco", lastName: "Torres Martinez", office: "Sierra Vista", position: "Laborer" },
  { id: "emp547", number: "30449", firstName: "Alvaro", lastName: "Tovar", office: "Tucson", position: "" },
  { id: "emp548", number: "2556", firstName: "Cole", lastName: "Townson", office: "Tucson", position: "Laborer" },
  { id: "emp549", number: "3045", firstName: "Anselmo", lastName: "Trejo Espinoza", office: "Tucson", position: "Pipelayer" },
  { id: "emp550", number: "2162", firstName: "Jonathan", lastName: "Trisler", office: "Tucson", position: "Operator" },
  { id: "emp551", number: "2983", firstName: "Levi", lastName: "Trisler", office: "Tucson", position: "Laborer" },
  { id: "emp552", number: "2171", firstName: "David", lastName: "Tucker", office: "Sierra Vista", position: "" },
  { id: "emp553", number: "2979", firstName: "Gene", lastName: "Tucker", office: "Tucson", position: "Operator" },
  { id: "emp554", number: "3012", firstName: "Thomas", lastName: "Tunks", office: "Tucson", position: "Operator" },
  { id: "emp555", number: "46181", firstName: "Daniel", lastName: "Uchimura", office: "Sierra Vista", position: "Shop Laborer" },
  { id: "emp556", number: "1469", firstName: "Jereme", lastName: "Ulmer", office: "Tucson", position: "Operator" },
  { id: "emp557", number: "2907", firstName: "Larry", lastName: "Underwood", office: "Tucson", position: "Pipelayer" },
  { id: "emp558", number: "2786", firstName: "Juanito", lastName: "Urias", office: "Tucson", position: "Operator" },
  { id: "emp559", number: "1165", firstName: "Greg", lastName: "Vacura", office: "Tucson", position: "Operator" },
  { id: "emp560", number: "2102", firstName: "Luis", lastName: "Valdez", office: "Tucson", position: "Operator" },
  { id: "emp561", number: "3029", firstName: "Edward", lastName: "Valencia", office: "Tucson", position: "Field Engineer" },
  { id: "emp562", number: "2931", firstName: "Francisco", lastName: "Valencia Medina", office: "Sierra Vista", position: "Laborer" },
  { id: "emp563", number: "12634", firstName: "Gregorio", lastName: "Valenzuela", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp564", number: "2938", firstName: "Robert", lastName: "Valenzuela", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp565", number: "3107", firstName: "Roberto", lastName: "Valenzuela", office: "Sierra Vista", position: "Operator" },
  { id: "emp566", number: "2323", firstName: "Sergio", lastName: "Valenzuela Hernandez", office: "Sierra Vista", position: "Concrete Finisher" },
  { id: "emp567", number: "3024", firstName: "Luis", lastName: "Valenzuela Navarro", office: "Tucson", position: "Laborer" },
  { id: "emp568", number: "2535", firstName: "Manuel", lastName: "Valenzuela Salazar", office: "Tucson", position: "Laborer" },
  { id: "emp569", number: "1959", firstName: "Pedro", lastName: "Valladarez Ruiz", office: "Tucson", position: "Project Administrator" },
  { id: "emp570", number: "2627", firstName: "Aurora", lastName: "Vandriel", office: "Tucson", position: "Field Engineer" },
  { id: "emp571", number: "2696", firstName: "Eathan", lastName: "Vandriel", office: "Tucson", position: "Laborer" },
  { id: "emp572", number: "3021", firstName: "Amanda", lastName: "Vanwinkle", office: "Tucson", position: "" },
  { id: "emp573", number: "2383", firstName: "Joaquin", lastName: "Vasquez", office: "Tucson", position: "Field Engineer" },
  { id: "emp574", number: "2954", firstName: "Jose", lastName: "Vasquez", office: "Sierra Vista", position: "Laborer" },
  { id: "emp575", number: "3108", firstName: "Rogelio", lastName: "Vasquez", office: "Sierra Vista", position: "Cdl Driver" },
  { id: "emp576", number: "2934", firstName: "Elian", lastName: "Vega", office: "Sierra Vista", position: "Mechanic" },
  { id: "emp577", number: "2614", firstName: "William", lastName: "Vega", office: "Tucson", position: "Operator" },
  { id: "emp578", number: "2093", firstName: "Salvador", lastName: "Velasquez", office: "Tucson", position: "Superintendent" },
  { id: "emp579", number: "2976", firstName: "Antonio", lastName: "Verdugo", office: "Tucson", position: "Superintendent" },
  { id: "emp580", number: "3123", firstName: "Brayan", lastName: "Vidana Velasquez", office: "Sierra Vista", position: "Laborer" },
  { id: "emp581", number: "2981", firstName: "Anthony", lastName: "Vigil", office: "Tucson", position: "Operator" },
  { id: "emp582", number: "3026", firstName: "Joseph", lastName: "Vigil", office: "Tucson", position: "Laborer" },
  { id: "emp583", number: "2685", firstName: "Victor", lastName: "Villa", office: "Tucson", position: "Laborer" },
  { id: "emp584", number: "2939", firstName: "Erik", lastName: "Villa Gonzalez", office: "Sierra Vista", position: "Field Engineer" },
  { id: "emp585", number: "2960", firstName: "Jesus", lastName: "Villa Juarez", office: "Sierra Vista", position: "Operator" },
  { id: "emp586", number: "2965", firstName: "David", lastName: "Villalovos", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp587", number: "2262", firstName: "Benito", lastName: "Virgen", office: "Tucson", position: "Pipelayer" },
  { id: "emp588", number: "2857", firstName: "Tanner", lastName: "Waltz", office: "Tucson", position: "Leadman" },
  { id: "emp589", number: "7045", firstName: "Devon", lastName: "Ward", office: "Sierra Vista", position: "Operator" },
  { id: "emp590", number: "2902", firstName: "Jacob", lastName: "Wash", office: "Tucson", position: "Laborer" },
  { id: "emp591", number: "3058", firstName: "James", lastName: "Wash", office: "Tucson", position: "Foreman" },
  { id: "emp592", number: "2157", firstName: "Kristine", lastName: "Wash", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp593", number: "1809", firstName: "Luke", lastName: "Weinstein", office: "Tucson", position: "Project Manager_Estimator" },
  { id: "emp594", number: "2993", firstName: "Devin", lastName: "Wheeler", office: "Tucson", position: "Operator" },
  { id: "emp595", number: "2882", firstName: "Nathan", lastName: "Wheeler", office: "Tucson", position: "Operator" },
  { id: "emp596", number: "3044", firstName: "Austin", lastName: "Williams", office: "Tucson", position: "" },
  { id: "emp597", number: "2634", firstName: "Justin", lastName: "Wilson", office: "Sierra Vista", position: "Division Manager" },
  { id: "emp598", number: "2143", firstName: "Robert", lastName: "Wilson", office: "Tucson", position: "Operator" },
  { id: "emp599", number: "1473", firstName: "Kevin", lastName: "Wilusz", office: "Sierra Vista", position: "Project Manager_Estimator" },
  { id: "emp600", number: "1739", firstName: "Patrick", lastName: "Wimpari", office: "Tucson", position: "" },
  { id: "emp601", number: "2144", firstName: "Tucker", lastName: "Wood", office: "Tucson", position: "Mechanic" },
  { id: "emp602", number: "1701", firstName: "Russell", lastName: "Woodman", office: "Tucson", position: "Superintendent" },
  { id: "emp603", number: "2720", firstName: "Tyler", lastName: "Yaryan", office: "Tucson", position: "Laborer" },
  { id: "emp604", number: "3059", firstName: "Benjamin", lastName: "Yepiz", office: "Tucson", position: "Operator" },
  { id: "emp605", number: "2443", firstName: "Francisco", lastName: "Young", office: "Sierra Vista", position: "Laborer" },
  { id: "emp606", number: "55759", firstName: "Nicholas", lastName: "Zedaker", office: "Tucson", position: "Foreman" },
  { id: "emp607", number: "50707", firstName: "Robert", lastName: "Zedaker", office: "Tucson", position: "Senior Director" },
  { id: "emp608", number: "3100", firstName: "Michael", lastName: "Zukowski", office: "Sierra Vista", position: "Laborer" },
  { id: "emp609", number: "3118", firstName: "Jordi", lastName: "Zuniga Samaniego", office: "Sierra Vista", position: "Laborer" },
  { id: "emp610", number: "6004", firstName: "Arnoldo", lastName: "Acosta", office: "Douglas (Maddux)", position: "Operator 1" },
  { id: "emp611", number: "6039", firstName: "Pablo", lastName: "Arreola", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp612", number: "6005", firstName: "Daniel", lastName: "Ballesteros", office: "Douglas (Maddux)", position: "Foreman" },
  { id: "emp613", number: "6031", firstName: "David", lastName: "Castaneda  Lopez", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp614", number: "6006", firstName: "Martin", lastName: "Cornejo", office: "Douglas (Maddux)", position: "Operator 1" },
  { id: "emp615", number: "6009", firstName: "Joaquin", lastName: "Escalante", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp616", number: "6010", firstName: "Omar", lastName: "Flores", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp617", number: "6011", firstName: "Jesus", lastName: "Garcia", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp618", number: "6007", firstName: "Brandon", lastName: "Garcia Cornidez", office: "Douglas (Maddux)", position: "Laborer 1" },
  { id: "emp619", number: "6008", firstName: "Bryan", lastName: "Garcia Cornidez", office: "Douglas (Maddux)", position: "Mechanic" },
  { id: "emp620", number: "6012", firstName: "Alex", lastName: "Gonzales", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp621", number: "6013", firstName: "Francisco", lastName: "Grijalva", office: "Douglas (Maddux)", position: "Operator 1" },
  { id: "emp622", number: "6015", firstName: "Emeterio", lastName: "Hernandez", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp623", number: "6016", firstName: "Jesus", lastName: "Hernandez", office: "Douglas (Maddux)", position: "Operator 1" },
  { id: "emp624", number: "6035", firstName: "Jesus", lastName: "Hernandez", office: "Douglas (Maddux)", position: "Laborer 1" },
  { id: "emp625", number: "6032", firstName: "Francisco", lastName: "Icedo", office: "Douglas (Maddux)", position: "Laborer 1" },
  { id: "emp626", number: "6017", firstName: "Stephany", lastName: "Icedo", office: "Douglas (Maddux)", position: "Admn/Pm/Pe/Fe/Frmn/Sups/Safety" },
  { id: "emp627", number: "6018", firstName: "Adolfo", lastName: "Lopez Lopez", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp628", number: "6019", firstName: "Saul", lastName: "Luna", office: "Douglas (Maddux)", position: "Laborer 1" },
  { id: "emp629", number: "6001", firstName: "Clinton", lastName: "Maddux", office: "Douglas (Maddux)", position: "Admn/Pm/Pe/Fe/Frmn/Sups/Safety" },
  { id: "emp630", number: "6002", firstName: "James", lastName: "Maddux", office: "Douglas (Maddux)", position: "Admn/Pm/Pe/Fe/Frmn/Sups/Safety" },
  { id: "emp631", number: "6020", firstName: "Ismael", lastName: "Mata", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp632", number: "6021", firstName: "Dennie", lastName: "Mcfarland", office: "Douglas (Maddux)", position: "Operator 1" },
  { id: "emp633", number: "6022", firstName: "Krystal", lastName: "Medrano", office: "Douglas (Maddux)", position: "Admn/Pm/Pe/Fe/Frmn/Sups/Safety" },
  { id: "emp634", number: "6023", firstName: "Oscar", lastName: "Mendez", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp635", number: "6024", firstName: "Enrique", lastName: "Pedrego", office: "Douglas (Maddux)", position: "Mechanic" },
  { id: "emp636", number: "6025", firstName: "Ramon", lastName: "Peralta", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp637", number: "6040", firstName: "Isiah", lastName: "Robles", office: "Douglas (Maddux)", position: "Laborer 1" },
  { id: "emp638", number: "6026", firstName: "Manuel", lastName: "Samaniego", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp639", number: "6033", firstName: "Jacinda", lastName: "Saucedo", office: "Douglas (Maddux)", position: "Admn/Pm/Pe/Fe/Frmn/Sups/Safety" },
  { id: "emp640", number: "6037", firstName: "Julian", lastName: "Solis Lianos", office: "Douglas (Maddux)", position: "Laborer 1" },
  { id: "emp641", number: "6027", firstName: "Roman", lastName: "Trujillo", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp642", number: "6036", firstName: "Robert", lastName: "Vance", office: "Douglas (Maddux)", position: "Operator 1" },
  { id: "emp643", number: "1010", firstName: "Erik", lastName: "Vega", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp644", number: "6030", firstName: "Gerardo", lastName: "Vincent", office: "Douglas (Maddux)", position: "Admn/Pm/Pe/Fe/Frmn/Sups/Safety" },
  { id: "emp645", number: "6034", firstName: "Jacob", lastName: "Vincent", office: "Douglas (Maddux)", position: "Operator 1" },
  { id: "emp646", number: "6041", firstName: "James", lastName: "O'Neil", office: "Douglas (Maddux)", position: "Cdl Driver" },
  { id: "emp647", number: "6042", firstName: "Saul G", lastName: "Luna", office: "Douglas (Maddux)", position: "Laborer 1" },
  { id: "emp648", number: "6043", firstName: "Manuel", lastName: "Martinez", office: "Douglas (Maddux)", position: "Cdl Driver" },
];

// ─── Storage keys ─────────────────────────────────────────────────────
// employees-roster        -> shared JSON array of {number, firstName, lastName, office}
// award-categories        -> shared JSON array of {id, name, winnerCount, nominees:[{id, firstName, lastName, office}]}
// admin-passcode          -> shared string
// votes:{employeeNumber}  -> shared JSON {employeeNumber, voterFirstName, voterLastName, office, timestamp, selections:{categoryId: nomineeId}}

const ROSTER_KEY = "employees-roster";
const CATEGORIES_KEY = "award-categories";
const PASSCODE_KEY = "admin-passcode";
const VOTE_PREFIX = "votes:";

async function safeGet(key, shared) {
  try {
    const res = await window.storage.get(key, shared);
    return res ? res.value : null;
  } catch (e) {
    return null;
  }
}
async function safeSet(key, value, shared) {
  try {
    return await window.storage.set(key, value, shared);
  } catch (e) {
    return null;
  }
}

const loadRoster = async () => {
  const v = await safeGet(ROSTER_KEY, true);
  return v ? JSON.parse(v) : [];
};
const saveRoster = (roster) => safeSet(ROSTER_KEY, JSON.stringify(roster), true);

const loadCategories = async () => {
  const v = await safeGet(CATEGORIES_KEY, true);
  return v ? JSON.parse(v) : [];
};
const saveCategories = (cats) => safeSet(CATEGORIES_KEY, JSON.stringify(cats), true);

const loadPasscode = async () => safeGet(PASSCODE_KEY, true);
const savePasscode = (code) => safeSet(PASSCODE_KEY, code, true);

const getVoteRecord = async (employeeNumber) => {
  const v = await safeGet(VOTE_PREFIX + employeeNumber, true);
  if (!v) return null;
  try {
    return JSON.parse(v);
  } catch (e) {
    return null;
  }
};
const hasVoted = async (employeeNumber) => !!(await getVoteRecord(employeeNumber));
const submitVoteRecord = (employeeNumber, record) =>
  safeSet(VOTE_PREFIX + employeeNumber, JSON.stringify(record), true);

const loadAllVotes = async () => {
  try {
    const listRes = await window.storage.list(VOTE_PREFIX, true);
    if (!listRes || !listRes.keys) return [];
    const results = [];
    for (const key of listRes.keys) {
      const v = await safeGet(key, true);
      if (v) {
        try {
          results.push(JSON.parse(v));
        } catch (e) {}
      }
    }
    return results;
  } catch (e) {
    return [];
  }
};

// ─── Utilities ────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2, 10);

const splitName = (fullName) => {
  const parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return { firstName: parts[0], lastName: "" };
  const lastName = parts[parts.length - 1];
  const firstName = parts.slice(0, -1).join(" ");
  return { firstName, lastName };
};

const fullName = (p) => `${p.firstName || ""} ${p.lastName || ""}`.trim();

const byLastName = (a, b) => {
  const la = (a.lastName || "").toLowerCase();
  const lb = (b.lastName || "").toLowerCase();
  if (la !== lb) return la < lb ? -1 : 1;
  const fa = (a.firstName || "").toLowerCase();
  const fb = (b.firstName || "").toLowerCase();
  return fa < fb ? -1 : fa > fb ? 1 : 0;
};

const normalize = (s) => (s || "").trim().toLowerCase().replace(/\s+/g, " ");

// Lightweight fuzzy match: exact substring match scores highest, otherwise
// checks that every query character appears in order somewhere in the text
// (typo/partial tolerant). Returns true/false plus a rough relevance score.
const fuzzyMatch = (query, text) => {
  const q = normalize(query);
  const t = normalize(text);
  if (!q) return { match: true, score: 0 };
  if (t.includes(q)) return { match: true, score: t.indexOf(q) === 0 ? 2 : 1 };
  let qi = 0;
  for (let ti = 0; ti < t.length && qi < q.length; ti++) {
    if (t[ti] === q[qi]) qi++;
  }
  return { match: qi === q.length, score: -1 };
};

function downloadJSON(filename, obj) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function downloadCSV(filename, rows) {
  const csv = rows
    .map((row) =>
      row
        .map((cell) => {
          const s = String(cell ?? "");
          if (s.includes(",") || s.includes('"') || s.includes("\n")) {
            return `"${s.replace(/"/g, '""')}"`;
          }
          return s;
        })
        .join(",")
    )
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("File read failed"));
    reader.readAsArrayBuffer(file);
  });
}

// Best-effort guess at which spreadsheet column maps to a given field,
// based on common header names. The admin can always override via dropdown.
const guessHeader = (headers, candidates) => {
  const norm = (s) => String(s).toLowerCase().trim();
  for (const h of headers) {
    if (candidates.includes(norm(h))) return h;
  }
  for (const h of headers) {
    if (candidates.some((c) => norm(h).includes(c))) return h;
  }
  return "";
};

// ─── Styles ───────────────────────────────────────────────────────────
const css = `
  @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@400;500;600&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  .eca-root { font-family: 'Barlow', sans-serif; background: ${COLORS.offWhite}; color: ${COLORS.navy}; min-height: 100%; }
  .eca-header {
    background: ${COLORS.navyDark}; padding: 20px 28px;
    display: flex; align-items: center; justify-content: space-between;
  }
  .eca-brand { font-family: 'Barlow Condensed', sans-serif; font-size: 22px; font-weight: 700; color: white; letter-spacing: 0.5px; }
  .eca-brand .accent { color: ${COLORS.orange}; }
  .eca-brand-sub { font-size: 11px; color: rgba(255,255,255,0.55); letter-spacing: 1.5px; text-transform: uppercase; margin-top: 2px; }
  .eca-nav-link { font-size: 12px; color: rgba(255,255,255,0.5); cursor: pointer; }
  .eca-nav-link:hover { color: white; }
  .eca-shell { max-width: 760px; margin: 0 auto; padding: 40px 20px 80px; }
  .eca-shell-wide { max-width: 1080px; margin: 0 auto; padding: 32px 24px 80px; }
  .eca-card {
    background: white; border-radius: 12px; border: 1px solid ${COLORS.gray200};
    padding: 28px 30px; box-shadow: 0 2px 10px rgba(13,31,60,0.04);
  }
  .eca-title { font-family: 'Barlow Condensed', sans-serif; font-size: 26px; font-weight: 700; color: ${COLORS.navy}; margin-bottom: 6px; }
  .eca-subtitle { font-size: 14px; color: ${COLORS.gray600}; margin-bottom: 22px; line-height: 1.5; }
  .eca-label { font-size: 12px; font-weight: 600; color: ${COLORS.gray600}; letter-spacing: 0.3px; display: block; margin-bottom: 5px; }
  .eca-input {
    width: 100%; padding: 11px 13px; border: 1px solid ${COLORS.gray200}; border-radius: 8px;
    font-family: 'Barlow', sans-serif; font-size: 14px; color: ${COLORS.navy}; background: white;
  }
  .eca-input:focus { outline: none; border-color: ${COLORS.blue}; }
  .eca-field { margin-bottom: 16px; }
  .eca-btn {
    display: inline-flex; align-items: center; justify-content: center; gap: 6px;
    padding: 11px 22px; border-radius: 8px; border: none;
    font-family: 'Barlow', sans-serif; font-size: 14px; font-weight: 600;
    cursor: pointer; transition: all 0.15s;
  }
  .eca-btn-primary { background: ${COLORS.orange}; color: white; }
  .eca-btn-primary:hover { background: ${COLORS.orangeDark}; }
  .eca-btn-primary:disabled { background: ${COLORS.gray200}; color: ${COLORS.gray400}; cursor: not-allowed; }
  .eca-btn-secondary { background: white; color: ${COLORS.navy}; border: 1px solid ${COLORS.gray200}; }
  .eca-btn-secondary:hover { background: ${COLORS.gray100}; }
  .eca-btn-sm { padding: 7px 14px; font-size: 12px; }
  .eca-btn-danger { background: ${COLORS.red}; color: white; }
  .eca-btn-block { width: 100%; }
  .eca-error { background: ${COLORS.redLight}; color: ${COLORS.red}; padding: 10px 14px; border-radius: 8px; font-size: 13px; margin-bottom: 16px; font-weight: 500; }
  .eca-success { background: ${COLORS.greenLight}; color: ${COLORS.green}; padding: 10px 14px; border-radius: 8px; font-size: 13px; margin-bottom: 16px; font-weight: 500; }
  .eca-category-block { border: 1px solid ${COLORS.gray200}; border-radius: 10px; margin-bottom: 14px; overflow: hidden; }
  .eca-category-head { padding: 14px 18px; background: ${COLORS.offWhite}; display: flex; align-items: center; justify-content: space-between; cursor: pointer; }
  .eca-category-name { font-family: 'Barlow Condensed', sans-serif; font-size: 17px; font-weight: 700; color: ${COLORS.navy}; }
  .eca-category-meta { font-size: 11px; color: ${COLORS.gray400}; margin-top: 2px; }
  .eca-chosen-pill { background: ${COLORS.greenLight}; color: ${COLORS.green}; font-size: 11px; font-weight: 700; padding: 3px 10px; border-radius: 100px; }
  .eca-skip-pill { background: ${COLORS.gray100}; color: ${COLORS.gray400}; font-size: 11px; font-weight: 600; padding: 3px 10px; border-radius: 100px; }
  .eca-nominee-list { padding: 6px 10px 14px; }
  .eca-nominee-row {
    display: flex; align-items: center; gap: 10px; padding: 10px 8px; border-radius: 8px; cursor: pointer;
  }
  .eca-nominee-row:hover { background: ${COLORS.offWhite}; }
  .eca-nominee-row.selected { background: rgba(232,98,26,0.08); }
  .eca-radio { width: 18px; height: 18px; border-radius: 50%; border: 2px solid ${COLORS.gray200}; flex-shrink: 0; display: flex; align-items: center; justify-content: center; }
  .eca-radio.selected { border-color: ${COLORS.orange}; }
  .eca-radio.selected::after { content: ''; width: 9px; height: 9px; border-radius: 50%; background: ${COLORS.orange}; }
  .eca-nominee-name { font-size: 14px; font-weight: 500; color: ${COLORS.navy}; }
  .eca-nominee-office { font-size: 11px; color: ${COLORS.gray400}; }
  .eca-progress-track { height: 6px; background: ${COLORS.gray200}; border-radius: 3px; overflow: hidden; margin-bottom: 24px; }
  .eca-progress-fill { height: 100%; background: ${COLORS.orange}; transition: width 0.25s; }
  .eca-stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 22px; }
  .eca-stat { background: white; border: 1px solid ${COLORS.gray200}; border-radius: 10px; padding: 16px 18px; }
  .eca-stat-label { font-size: 10px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: ${COLORS.gray400}; margin-bottom: 6px; }
  .eca-stat-value { font-family: 'Barlow Condensed', sans-serif; font-size: 28px; font-weight: 700; color: ${COLORS.navy}; }
  .eca-tabs { display: flex; gap: 4px; margin-bottom: 22px; border-bottom: 1px solid ${COLORS.gray200}; }
  .eca-tab { padding: 10px 16px; font-size: 13px; font-weight: 600; color: ${COLORS.gray400}; cursor: pointer; border-bottom: 2px solid transparent; }
  .eca-tab.active { color: ${COLORS.navy}; border-bottom-color: ${COLORS.orange}; }
  .eca-table-wrap { overflow-x: auto; }
  table.eca-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  .eca-table th { text-align: left; padding: 9px 12px; font-size: 10px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; color: ${COLORS.gray400}; background: ${COLORS.gray100}; }
  .eca-table td { padding: 10px 12px; border-bottom: 1px solid ${COLORS.gray100}; }
  .eca-tag { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 500; background: ${COLORS.gray100}; color: ${COLORS.gray600}; }
  .eca-empty { text-align: center; padding: 36px 20px; color: ${COLORS.gray400}; font-size: 13px; }
  .eca-row-flex { display: flex; gap: 10px; align-items: center; }
  .eca-mini-btn { background: transparent; border: none; color: ${COLORS.gray400}; cursor: pointer; font-size: 12px; padding: 4px 8px; }
  .eca-mini-btn:hover { color: ${COLORS.red}; }
  .eca-section-title { font-family: 'Barlow Condensed', sans-serif; font-size: 16px; font-weight: 700; color: ${COLORS.navy}; margin-bottom: 12px; margin-top: 4px; }
  .eca-divider { height: 1px; background: ${COLORS.gray200}; margin: 24px 0; }
  textarea.eca-textarea { width: 100%; min-height: 110px; padding: 11px 13px; border: 1px solid ${COLORS.gray200}; border-radius: 8px; font-family: 'Barlow', sans-serif; font-size: 13px; resize: vertical; }
  .eca-help { font-size: 11px; color: ${COLORS.gray400}; margin-top: 4px; line-height: 1.5; }
  .eca-bar-track { flex: 1; height: 8px; background: ${COLORS.gray100}; border-radius: 4px; overflow: hidden; }
  .eca-bar-fill { height: 100%; background: ${COLORS.blue}; border-radius: 4px; }
`;

// ─── Loading shell ────────────────────────────────────────────────────
function LoadingScreen() {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "60vh", flexDirection: "column", gap: 10 }}>
      <div style={{ fontFamily: "'Barlow Condensed', sans-serif", fontSize: 20, fontWeight: 700, color: COLORS.navy }}>Loading…</div>
      <div style={{ fontSize: 12, color: COLORS.gray400 }}>Connecting to the awards system</div>
    </div>
  );
}

// ─── Header ───────────────────────────────────────────────────────────
function Header({ mode, onSwitch }) {
  return (
    <div className="eca-header">
      <div>
        <div className="eca-brand">KE<span className="accent">&amp;</span>G</div>
        <div className="eca-brand-sub">Employee Choice Awards</div>
      </div>
      <div className="eca-nav-link" onClick={onSwitch}>
        {mode === "admin" ? "← Back to voting" : "Admin"}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════
// VOTER FLOW
// ══════════════════════════════════════════════════════════════════════

function VoterLogin({ roster, onVerified }) {
  const [empNumber, setEmpNumber] = useState("");
  const [name, setName] = useState("");
  const [error, setError] = useState("");
  const [checking, setChecking] = useState(false);

  const submit = async () => {
    setError("");
    const num = empNumber.trim();
    const typedName = normalize(name);
    if (!num || !name.trim()) {
      setError("Enter your employee number and your name.");
      return;
    }
    const emp = roster.find((e) => String(e.number).trim() === num);
    if (!emp) {
      setError("We couldn't find that employee number. Check the number and try again, or contact HR.");
      return;
    }
    const rosterFull = normalize(fullName(emp));
    const rosterReversed = normalize(`${emp.lastName || ""} ${emp.firstName || ""}`);
    if (typedName !== rosterFull && typedName !== rosterReversed) {
      setError("That name doesn't match the employee number on file. Check your spelling and try again, or contact HR.");
      return;
    }
    setChecking(true);
    const existingRecord = await getVoteRecord(emp.number);
    setChecking(false);
    onVerified(emp, existingRecord);
  };

  return (
    <div className="eca-shell">
      <div className="eca-card">
        <div className="eca-title">Cast Your Ballot</div>
        <div className="eca-subtitle">
          Enter your employee number and your full name to verify you're eligible to vote. You can skip any category you don't want to vote in, and if you've already voted, you can log back in to see your picks or vote in any category you haven't gotten to yet.
        </div>
        {error && <div className="eca-error">{error}</div>}
        <div className="eca-field">
          <label className="eca-label">Employee Number</label>
          <input className="eca-input" value={empNumber} onChange={(e) => setEmpNumber(e.target.value)} placeholder="e.g. 10234" />
        </div>
        <div className="eca-field">
          <label className="eca-label">Full Name</label>
          <input className="eca-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="First Last" onKeyDown={(e) => e.key === "Enter" && submit()} />
        </div>
        <button className="eca-btn eca-btn-primary eca-btn-block" onClick={submit} disabled={checking}>
          {checking ? "Checking…" : "Continue to Ballot"}
        </button>
      </div>
    </div>
  );
}

function VotingForm({ voter, categories, existingRecord, onSubmitted }) {
  const lockedSelections = existingRecord?.selections || {};
  const lockedCatIds = new Set(Object.keys(lockedSelections));
  const isReturningVoter = !!existingRecord;
  const openCategories = categories.filter((cat) => !lockedCatIds.has(cat.id));

  const [selections, setSelections] = useState({});
  const [openCat, setOpenCat] = useState(openCategories[0]?.id || null);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");
  const [nomineeSearch, setNomineeSearch] = useState({});

  const newVotedCount = Object.keys(selections).length;
  const allDone = categories.length > 0 && lockedCatIds.size === categories.length;

  const choose = (catId, nomineeId) => {
    if (lockedCatIds.has(catId)) return; // already voted in this category, can't change
    setSelections((s) => {
      const next = { ...s };
      if (next[catId] === nomineeId) {
        delete next[catId]; // click again to deselect
      } else {
        next[catId] = nomineeId;
      }
      return next;
    });
  };

  const submit = async () => {
    setError("");
    if (!isReturningVoter && newVotedCount === 0) {
      setError("You haven't selected anyone yet. Choose at least one category, or contact HR if you'd like to abstain entirely.");
      return;
    }
    if (isReturningVoter && newVotedCount === 0) {
      setError("Pick someone in at least one of your remaining categories, or come back later — your existing votes are already saved.");
      return;
    }
    setSubmitting(true);
    const mergedSelections = { ...lockedSelections, ...selections };
    const record = {
      employeeNumber: voter.number,
      voterFirstName: voter.firstName,
      voterLastName: voter.lastName,
      office: voter.office || "",
      timestamp: existingRecord?.timestamp || new Date().toISOString(),
      lastUpdated: new Date().toISOString(),
      selections: mergedSelections,
    };
    await submitVoteRecord(voter.number, record);
    setSubmitting(false);
    onSubmitted();
  };

  return (
    <div className="eca-shell">
      <div style={{ marginBottom: 18 }}>
        <div className="eca-title" style={{ fontSize: 22, marginBottom: 2 }}>
          {isReturningVoter ? `Welcome back, ${voter.firstName}` : `Hi ${voter.firstName}, cast your votes below`}
        </div>
        <div className="eca-subtitle" style={{ marginBottom: 0 }}>
          {isReturningVoter
            ? allDone
              ? "You've already voted in every category — here's what you picked. Nothing left to do."
              : `You've already voted in ${lockedCatIds.size} of ${categories.length} categories. Those are locked in below. Anything left open you can still vote in.`
            : `${newVotedCount} of ${categories.length} categories selected. You can leave categories blank — nothing is required.`}
        </div>
      </div>
      <div className="eca-progress-track">
        <div className="eca-progress-fill" style={{ width: `${categories.length ? ((lockedCatIds.size + newVotedCount) / categories.length) * 100 : 0}%` }} />
      </div>

      {error && <div className="eca-error">{error}</div>}

      {categories.map((cat) => {
        const locked = lockedCatIds.has(cat.id);
        const isOpen = openCat === cat.id;
        const chosenId = locked ? lockedSelections[cat.id] : selections[cat.id];
        const chosenNominee = cat.nominees.find((n) => n.id === chosenId);
        return (
          <div key={cat.id} className="eca-category-block">
            <div
              className="eca-category-head"
              style={locked ? { cursor: "default" } : undefined}
              onClick={() => !locked && setOpenCat(isOpen ? null : cat.id)}
            >
              <div>
                <div className="eca-category-name">{cat.name}</div>
                <div className="eca-category-meta">{cat.nominees.length} nominees{locked ? " · already voted" : ""}</div>
              </div>
              {chosenNominee ? (
                <span className="eca-chosen-pill">✓ {locked ? "You voted: " : ""}{fullName(chosenNominee)}</span>
              ) : locked ? (
                <span className="eca-skip-pill">Skipped</span>
              ) : (
                <span className="eca-skip-pill">Not voted</span>
              )}
            </div>
            {isOpen && !locked && (
              <div className="eca-nominee-list">
                {cat.nominees.length > 6 && (
                  <input
                    className="eca-input"
                    style={{ marginBottom: 10 }}
                    placeholder="Search by name…"
                    value={nomineeSearch[cat.id] || ""}
                    onChange={(e) => setNomineeSearch((s) => ({ ...s, [cat.id]: e.target.value }))}
                    onClick={(e) => e.stopPropagation()}
                  />
                )}
                {cat.nominees
                  .map((n) => ({ n, m: fuzzyMatch(nomineeSearch[cat.id] || "", fullName(n)) }))
                  .filter((x) => x.m.match)
                  .sort((a, b) => (b.m.score - a.m.score) || byLastName(a.n, b.n))
                  .map(({ n }) => (
                    <div
                      key={n.id}
                      className={`eca-nominee-row${chosenId === n.id ? " selected" : ""}`}
                      onClick={() => choose(cat.id, n.id)}
                    >
                      <div className={`eca-radio${chosenId === n.id ? " selected" : ""}`} />
                      <div>
                        <div className="eca-nominee-name">{fullName(n)}</div>
                        {n.office && <div className="eca-nominee-office">{n.office}</div>}
                      </div>
                    </div>
                  ))}
                {cat.nominees.length === 0 && <div className="eca-empty">No nominees loaded for this category yet.</div>}
                {cat.nominees.length > 0 && (nomineeSearch[cat.id] || "").trim() &&
                  cat.nominees.every((n) => !fuzzyMatch(nomineeSearch[cat.id], fullName(n)).match) && (
                    <div className="eca-empty">No one matches "{nomineeSearch[cat.id]}".</div>
                  )}
              </div>
            )}
          </div>
        );
      })}

      <div className="eca-divider" />
      {allDone ? (
        <div className="eca-help" style={{ textAlign: "center" }}>
          Thanks for taking part in the Employee Choice Awards.
        </div>
      ) : (
        <>
          <button className="eca-btn eca-btn-primary eca-btn-block" onClick={submit} disabled={submitting}>
            {submitting ? "Saving…" : isReturningVoter ? "Save My Remaining Votes" : "Submit My Ballot"}
          </button>
          <div className="eca-help" style={{ textAlign: "center", marginTop: 10 }}>
            {isReturningVoter
              ? "Categories you've already voted in are locked and can't be changed."
              : "You can log back in later with your employee number to vote in any category you skip."}
          </div>
        </>
      )}
    </div>
  );
}

function VoterDone() {
  return (
    <div className="eca-shell">
      <div className="eca-card" style={{ textAlign: "center", padding: "48px 30px" }}>
        <div style={{ fontSize: 40, marginBottom: 12 }}>✓</div>
        <div className="eca-title">Ballot Submitted</div>
        <div className="eca-subtitle" style={{ marginBottom: 0 }}>
          Thanks for taking part in the Employee Choice Awards. Winners will be announced at the Shareholders Dinner.
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════
// ADMIN FLOW
// ══════════════════════════════════════════════════════════════════════

function AdminLogin({ storedPasscode, onSetPasscode, onLoggedIn }) {
  const [code, setCode] = useState("");
  const [confirmCode, setConfirmCode] = useState("");
  const [error, setError] = useState("");

  const firstTime = storedPasscode === null || storedPasscode === "";

  const submit = async () => {
    setError("");
    if (firstTime) {
      if (code.length < 4) {
        setError("Choose a passcode at least 4 characters long.");
        return;
      }
      if (code !== confirmCode) {
        setError("Passcodes don't match.");
        return;
      }
      await onSetPasscode(code);
      onLoggedIn();
    } else {
      if (code !== storedPasscode) {
        setError("Incorrect passcode.");
        return;
      }
      onLoggedIn();
    }
  };

  return (
    <div className="eca-shell">
      <div className="eca-card">
        <div className="eca-title">{firstTime ? "Set Up Admin Access" : "Admin Login"}</div>
        <div className="eca-subtitle">
          {firstTime
            ? "This is the first time the admin panel has been opened. Choose a passcode you'll use to manage employees, categories, and results."
            : "Enter the admin passcode to manage the roster, categories, and results."}
        </div>
        {error && <div className="eca-error">{error}</div>}
        <div className="eca-field">
          <label className="eca-label">Passcode</label>
          <input className="eca-input" type="password" value={code} onChange={(e) => setCode(e.target.value)} onKeyDown={(e) => e.key === "Enter" && !firstTime && submit()} />
        </div>
        {firstTime && (
          <div className="eca-field">
            <label className="eca-label">Confirm Passcode</label>
            <input className="eca-input" type="password" value={confirmCode} onChange={(e) => setConfirmCode(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
          </div>
        )}
        <button className="eca-btn eca-btn-primary eca-btn-block" onClick={submit}>
          {firstTime ? "Set Passcode & Continue" : "Log In"}
        </button>
        <div className="eca-help" style={{ marginTop: 10 }}>
          This passcode is a convenience lock to keep casual visitors out of admin tools — it isn't encrypted security. Don't use it for anything beyond this event.
        </div>
      </div>
    </div>
  );
}

function AdminOverview({ roster, categories, votes }) {
  const votedCount = votes.length;
  const pct = roster.length ? Math.round((votedCount / roster.length) * 100) : 0;
  const byOffice = OFFICES.map((office) => ({
    office,
    total: roster.filter((e) => e.office === office).length,
    voted: votes.filter((v) => v.office === office).length,
  }));
  const [expandedLeaderboard, setExpandedLeaderboard] = useState(null);

  const standingsFor = (cat) => {
    const tally = {};
    votes.forEach((v) => {
      const choice = v.selections && v.selections[cat.id];
      if (choice) tally[choice] = (tally[choice] || 0) + 1;
    });
    return cat.nominees
      .map((n) => ({ nominee: n, count: tally[n.id] || 0 }))
      .sort((a, b) => b.count - a.count || byLastName(a.nominee, b.nominee));
  };

  return (
    <div>
      <div className="eca-stats-grid">
        <div className="eca-stat">
          <div className="eca-stat-label">Eligible Employees</div>
          <div className="eca-stat-value">{roster.length}</div>
        </div>
        <div className="eca-stat">
          <div className="eca-stat-label">Ballots Submitted</div>
          <div className="eca-stat-value">{votedCount}</div>
        </div>
        <div className="eca-stat">
          <div className="eca-stat-label">Participation</div>
          <div className="eca-stat-value">{pct}%</div>
        </div>
        <div className="eca-stat">
          <div className="eca-stat-label">Categories</div>
          <div className="eca-stat-value">{categories.length}</div>
        </div>
      </div>

      <div className="eca-card" style={{ marginBottom: 20 }}>
        <div className="eca-section-title">Participation by Office</div>
        {byOffice.map((o) => (
          <div key={o.office} style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
            <div style={{ width: 140, fontSize: 13, color: COLORS.gray600, flexShrink: 0 }}>{o.office}</div>
            <div className="eca-bar-track">
              <div className="eca-bar-fill" style={{ width: `${o.total ? (o.voted / o.total) * 100 : 0}%` }} />
            </div>
            <div style={{ width: 90, fontSize: 12, color: COLORS.gray400, textAlign: "right" }}>{o.voted} / {o.total}</div>
          </div>
        ))}
        {roster.some((e) => !e.office) && (
          <div className="eca-help">Some employees have no office assigned and aren't reflected above.</div>
        )}
      </div>

      <div className="eca-card" style={{ marginBottom: 20 }}>
        <div className="eca-section-title">Who's Winning, by Category</div>
        <div className="eca-help" style={{ marginBottom: 14 }}>
          Live standings from ballots submitted so far. The top vote-getters (equal to each category's winner count) are highlighted as currently leading — click a category to see full standings.
        </div>
        {categories.length === 0 && <div className="eca-empty">No categories set up yet.</div>}
        {categories.map((cat) => {
          const standings = standingsFor(cat);
          const leader = standings[0];
          const isExpanded = expandedLeaderboard === cat.id;
          const totalCatVotes = standings.reduce((s, x) => s + x.count, 0);
          return (
            <div key={cat.id} className="eca-category-block">
              <div className="eca-category-head" onClick={() => setExpandedLeaderboard(isExpanded ? null : cat.id)}>
                <div>
                  <div className="eca-category-name">{cat.name}</div>
                  <div className="eca-category-meta">{totalCatVotes} vote{totalCatVotes === 1 ? "" : "s"} cast · {cat.winnerCount} winner{cat.winnerCount === 1 ? "" : "s"}</div>
                </div>
                {leader && leader.count > 0 ? (
                  <span className="eca-chosen-pill">Leading: {fullName(leader.nominee)} ({leader.count})</span>
                ) : (
                  <span className="eca-skip-pill">No votes yet</span>
                )}
              </div>
              {isExpanded && (
                <div style={{ padding: "6px 18px 16px" }}>
                  {standings.length === 0 && <div className="eca-empty">No nominees in this category yet.</div>}
                  {standings.map((s, i) => (
                    <div key={s.nominee.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "7px 0", borderBottom: i < standings.length - 1 ? `1px solid ${COLORS.gray100}` : "none" }}>
                      <div style={{ width: 22, fontSize: 12, fontWeight: 700, color: i < cat.winnerCount ? COLORS.orange : COLORS.gray400 }}>{i + 1}</div>
                      <div style={{ flex: 1, fontSize: 13, fontWeight: i < cat.winnerCount ? 600 : 400, color: COLORS.navy }}>
                        {fullName(s.nominee)}{s.nominee.office ? <span style={{ color: COLORS.gray400, fontWeight: 400 }}> — {s.nominee.office}</span> : null}
                      </div>
                      <div className="eca-bar-track" style={{ maxWidth: 140 }}>
                        <div className="eca-bar-fill" style={{ width: `${totalCatVotes ? (s.count / totalCatVotes) * 100 : 0}%`, background: i < cat.winnerCount ? COLORS.orange : COLORS.blue }} />
                      </div>
                      <div style={{ width: 30, textAlign: "right", fontSize: 12, fontWeight: 600, color: COLORS.navy }}>{s.count}</div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div className="eca-card">
        <div className="eca-section-title">Votes per Category</div>
        {categories.length === 0 && <div className="eca-empty">No categories set up yet.</div>}
        {categories.map((cat) => {
          const catVotes = votes.filter((v) => v.selections && v.selections[cat.id]).length;
          return (
            <div key={cat.id} style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
              <div style={{ width: 220, fontSize: 13, color: COLORS.gray600, flexShrink: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{cat.name}</div>
              <div className="eca-bar-track">
                <div className="eca-bar-fill" style={{ width: `${votedCount ? (catVotes / votedCount) * 100 : 0}%`, background: COLORS.orange }} />
              </div>
              <div style={{ width: 40, fontSize: 12, color: COLORS.gray400, textAlign: "right" }}>{catVotes}</div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function AdminImport({ roster, setRoster, categories, setCategories, votes, onVotesReset }) {
  const [fileName, setFileName] = useState("");
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [mapping, setMapping] = useState({ number: "", nameMode: "split", fullName: "", firstName: "", lastName: "", office: "", position: "" });
  const [error, setError] = useState("");
  const [confirmReplace, setConfirmReplace] = useState(false);
  const [summary, setSummary] = useState(null);
  const [confirmWipe, setConfirmWipe] = useState(false);
  const [wipeMsg, setWipeMsg] = useState("");

  const loadSheet = (wb, sheetName) => {
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
    if (json.length === 0) {
      setHeaders([]);
      setRows([]);
      setError("That sheet doesn't have any data rows.");
      return;
    }
    const hdrs = Object.keys(json[0]);
    setHeaders(hdrs);
    setRows(json);
    setError("");
    setSummary(null);
    setMapping({
      number: guessHeader(hdrs, ["employee", "emp #", "employee number", "ee #", "number", "id", "employee id"]),
      nameMode: guessHeader(hdrs, ["first", "firstname", "first name"]) ? "split" : "combined",
      fullName: guessHeader(hdrs, ["name", "full name", "employee name"]),
      firstName: guessHeader(hdrs, ["first", "firstname", "first name"]),
      lastName: guessHeader(hdrs, ["last", "lastname", "last name"]),
      office: guessHeader(hdrs, ["office", "location", "udlocation", "region"]),
      position: guessHeader(hdrs, ["position", "positiondescription", "job title", "title", "position description"]),
    });
  };

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError("");
    setSummary(null);
    setFileName(file.name);
    try {
      const wb = await readExcelFile(file);
      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
      const first = wb.SheetNames[0];
      setSelectedSheet(first);
      loadSheet(wb, first);
    } catch (err) {
      setError("Couldn't read that file. Make sure it's a valid .xlsx or .xls file.");
    }
  };

  const changeSheet = (name) => {
    setSelectedSheet(name);
    if (workbook) loadSheet(workbook, name);
  };

  const setMap = (k, v) => setMapping((m) => ({ ...m, [k]: v }));

  const buildEmployee = (row) => {
    const number = String(row[mapping.number] ?? "").trim();
    let firstName = "", lastName = "";
    if (mapping.nameMode === "combined") {
      const split = splitName(String(row[mapping.fullName] ?? ""));
      firstName = split.firstName;
      lastName = split.lastName;
    } else {
      firstName = String(row[mapping.firstName] ?? "").trim();
      lastName = String(row[mapping.lastName] ?? "").trim();
    }
    const office = mapping.office ? String(row[mapping.office] ?? "").trim() : "";
    const position = mapping.position ? String(row[mapping.position] ?? "").trim() : "";
    return { number, firstName, lastName, office, position };
  };

  const mappingComplete = !!mapping.number && (
    (mapping.nameMode === "combined" && !!mapping.fullName) ||
    (mapping.nameMode === "split" && !!mapping.firstName && !!mapping.lastName)
  );

  const preview = mappingComplete ? rows.slice(0, 5).map(buildEmployee) : [];

  const commitImport = (mode) => {
    const mapped = rows.map(buildEmployee).filter((e) => e.number);
    const withIds = mapped.map((e) => ({ id: uid(), ...e }));
    let next;
    if (mode === "replace") {
      next = withIds;
    } else {
      const existingNumbers = new Set(roster.map((e) => e.number));
      const newOnes = withIds.filter((e) => !existingNumbers.has(e.number));
      next = [...roster, ...newOnes];
    }
    setRoster(next);
    saveRoster(next);
    setSummary({ imported: withIds.length, mode, total: next.length });
    setConfirmReplace(false);
  };

  const clearEverything = async () => {
    await saveRoster([]);
    await saveCategories([]);
    for (const v of votes) {
      await safeSet(VOTE_PREFIX + v.employeeNumber, "", true);
    }
    setRoster([]);
    setCategories([]);
    onVotesReset();
    setConfirmWipe(false);
    setWipeMsg(`Cleared — ${roster.length} employees, ${categories.length} categories, and ${votes.length} votes are gone. Upload a file below to start fresh.`);
    setFileName(""); setWorkbook(null); setHeaders([]); setRows([]); setSummary(null);
  };

  return (
    <div>
      <div className="eca-card" style={{ marginBottom: 18, borderColor: COLORS.red }}>
        <div className="eca-section-title" style={{ color: COLORS.red }}>Start Fresh</div>
        <div className="eca-help" style={{ marginBottom: 12 }}>
          Clears the current employee roster, all categories and nominees, and all submitted votes — everything in the app right now. Use this before importing a clean file for a new event. This can't be undone — grab a backup first from Settings → Backup &amp; Restore if there's anything here worth keeping.
        </div>
        {wipeMsg && <div className="eca-success">{wipeMsg}</div>}
        {!confirmWipe ? (
          <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={() => setConfirmWipe(true)}>Clear Everything</button>
        ) : (
          <div>
            <div className="eca-error">This deletes {roster.length} employees, {categories.length} categories, and {votes.length} votes. This can't be undone.</div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={clearEverything}>Yes, clear everything</button>
              <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmWipe(false)}>Cancel</button>
            </div>
          </div>
        )}
      </div>

      <div className="eca-card" style={{ marginBottom: 18 }}>
        <div className="eca-section-title">Import Employees From Excel</div>
        <div className="eca-help" style={{ marginBottom: 12 }}>
          Upload an .xlsx or .xls file. Any column layout works — you'll map your spreadsheet's columns to the fields the app uses on the next step.
        </div>
        {error && <div className="eca-error">{error}</div>}
        <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ fontSize: 13 }} />
        {fileName && <div className="eca-help" style={{ marginTop: 8 }}>Loaded: {fileName} ({rows.length} rows)</div>}

        {sheetNames.length > 1 && (
          <div className="eca-field" style={{ marginTop: 14, maxWidth: 300 }}>
            <label className="eca-label">Sheet</label>
            <select className="eca-input" value={selectedSheet} onChange={(e) => changeSheet(e.target.value)}>
              {sheetNames.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        )}

        {headers.length > 0 && (
          <>
            <div className="eca-divider" />
            <div className="eca-section-title" style={{ fontSize: 14 }}>Map Columns</div>
            <div className="eca-field">
              <label className="eca-label">Employee Number *</label>
              <select className="eca-input" value={mapping.number} onChange={(e) => setMap("number", e.target.value)}>
                <option value="">— Select column —</option>
                {headers.map((h) => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>

            <div className="eca-field">
              <label className="eca-label">Name Format</label>
              <select className="eca-input" value={mapping.nameMode} onChange={(e) => setMap("nameMode", e.target.value)}>
                <option value="split">First and Last are separate columns</option>
                <option value="combined">Full name is one column</option>
              </select>
            </div>

            {mapping.nameMode === "combined" ? (
              <div className="eca-field">
                <label className="eca-label">Full Name Column *</label>
                <select className="eca-input" value={mapping.fullName} onChange={(e) => setMap("fullName", e.target.value)}>
                  <option value="">— Select column —</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ) : (
              <div style={{ display: "flex", gap: 16 }}>
                <div className="eca-field" style={{ flex: 1 }}>
                  <label className="eca-label">First Name Column *</label>
                  <select className="eca-input" value={mapping.firstName} onChange={(e) => setMap("firstName", e.target.value)}>
                    <option value="">— Select column —</option>
                    {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
                <div className="eca-field" style={{ flex: 1 }}>
                  <label className="eca-label">Last Name Column *</label>
                  <select className="eca-input" value={mapping.lastName} onChange={(e) => setMap("lastName", e.target.value)}>
                    <option value="">— Select column —</option>
                    {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
              </div>
            )}

            <div style={{ display: "flex", gap: 16 }}>
              <div className="eca-field" style={{ flex: 1 }}>
                <label className="eca-label">Office Column (optional)</label>
                <select className="eca-input" value={mapping.office} onChange={(e) => setMap("office", e.target.value)}>
                  <option value="">— None —</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div className="eca-field" style={{ flex: 1 }}>
                <label className="eca-label">Position Column (optional)</label>
                <select className="eca-input" value={mapping.position} onChange={(e) => setMap("position", e.target.value)}>
                  <option value="">— None —</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>

            {mappingComplete && (
              <>
                <div className="eca-divider" />
                <div className="eca-section-title" style={{ fontSize: 14 }}>Preview (first 5 rows)</div>
                <div className="eca-table-wrap" style={{ marginBottom: 16 }}>
                  <table className="eca-table">
                    <thead><tr><th>Number</th><th>First</th><th>Last</th><th>Office</th><th>Position</th></tr></thead>
                    <tbody>
                      {preview.map((p, i) => (
                        <tr key={i}>
                          <td>{p.number}</td>
                          <td>{p.firstName}</td>
                          <td>{p.lastName}</td>
                          <td>{p.office}</td>
                          <td>{p.position}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {summary && (
                  <div className="eca-success">
                    Imported {summary.imported} employee{summary.imported === 1 ? "" : "s"} ({summary.mode === "replace" ? "replaced roster" : "added to roster"}). Roster now has {summary.total} employees.
                  </div>
                )}

                {!confirmReplace ? (
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <button className="eca-btn eca-btn-primary eca-btn-sm" onClick={() => setConfirmReplace(true)}>Replace Entire Roster ({rows.length})</button>
                    <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => commitImport("append")}>Add to Existing Roster</button>
                  </div>
                ) : (
                  <div className="eca-error">
                    This replaces all {roster.length} current employees with the {rows.length} rows from this file. This can't be undone.
                    <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                      <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={() => commitImport("replace")}>Yes, replace roster</button>
                      <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmReplace(false)}>Cancel</button>
                    </div>
                  </div>
                )}
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function AdminEmployees({ roster, setRoster, saving }) {
  const [modal, setModal] = useState(null); // null | 'new' | employee object
  const [bulkText, setBulkText] = useState("");
  const [showBulk, setShowBulk] = useState(false);
  const [search, setSearch] = useState("");
  const [confirmLoadOfficial, setConfirmLoadOfficial] = useState(false);

  const filtered = roster
    .filter((e) => {
      const q = search.toLowerCase();
      return !q || fullName(e).toLowerCase().includes(q) || String(e.number).includes(q) || (e.position || "").toLowerCase().includes(q);
    })
    .slice()
    .sort(byLastName);

  const persist = (next) => {
    setRoster(next);
    saveRoster(next);
  };

  const addOrUpdate = (form) => {
    if (modal === "new") {
      persist([...roster, { ...form, id: uid() }]);
    } else {
      persist(roster.map((e) => (e.id === form.id ? form : e)));
    }
    setModal(null);
  };

  const remove = (id) => {
    persist(roster.filter((e) => e.id !== id));
  };

  const loadOfficialRoster = () => {
    persist(OFFICIAL_ROSTER);
    setConfirmLoadOfficial(false);
  };

  const runBulkImport = () => {
    const lines = bulkText.split("\n").map((l) => l.trim()).filter(Boolean);
    const additions = [];
    for (const line of lines) {
      const parts = line.split(",").map((p) => p.trim());
      if (parts.length < 2) continue;
      const [number, name, office, position] = parts;
      const { firstName, lastName } = splitName(name);
      additions.push({ id: uid(), number, firstName, lastName, office: office || "", position: position || "" });
    }
    persist([...roster, ...additions]);
    setBulkText("");
    setShowBulk(false);
  };

  return (
    <div>
      {modal && (
        <EmployeeModal
          initial={modal === "new" ? {} : modal}
          onSave={addOrUpdate}
          onClose={() => setModal(null)}
        />
      )}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, gap: 10, flexWrap: "wrap" }}>
        <input className="eca-input" style={{ maxWidth: 260 }} placeholder="Search employees…" value={search} onChange={(e) => setSearch(e.target.value)} />
        <div className="eca-row-flex">
          <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmLoadOfficial(true)}>Load Official Roster ({OFFICIAL_ROSTER.length})</button>
          <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setShowBulk((s) => !s)}>Bulk Import</button>
          <button className="eca-btn eca-btn-primary eca-btn-sm" onClick={() => setModal("new")}>+ Add Employee</button>
        </div>
      </div>

      {confirmLoadOfficial && (
        <div className="eca-card" style={{ marginBottom: 16 }}>
          <div className="eca-error" style={{ marginBottom: 12 }}>
            This replaces the current employee list ({roster.length}) with the official roster ({OFFICIAL_ROSTER.length}: 610 KE&amp;G + 39 Maddux &amp; Sons). Anyone who has already voted stays voted — this only changes who's eligible going forward. This can't be undone.
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={loadOfficialRoster}>Yes, load official roster</button>
            <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmLoadOfficial(false)}>Cancel</button>
          </div>
        </div>
      )}

      {showBulk && (
        <div className="eca-card" style={{ marginBottom: 16 }}>
          <div className="eca-section-title">Bulk Import Employees</div>
          <div className="eca-help" style={{ marginBottom: 8 }}>
            One employee per line: <strong>employee number, full name, office, position</strong>. Office and position are optional. Example:<br />
            10234, John Smith, Tucson, Operator<br />
            10891, Maria Garcia, Sierra Vista, Foreman
          </div>
          <textarea className="eca-textarea" value={bulkText} onChange={(e) => setBulkText(e.target.value)} placeholder="10234, John Smith, Tucson, Operator" />
          <div style={{ marginTop: 10, display: "flex", gap: 8 }}>
            <button className="eca-btn eca-btn-primary eca-btn-sm" onClick={runBulkImport}>Import</button>
            <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setShowBulk(false)}>Cancel</button>
          </div>
        </div>
      )}

      <div className="eca-card" style={{ padding: 0 }}>
        <div className="eca-table-wrap">
          <table className="eca-table">
            <thead><tr><th>Emp #</th><th>Last</th><th>First</th><th>Office</th><th>Position</th><th></th></tr></thead>
            <tbody>
              {filtered.length === 0 && <tr><td colSpan={6}><div className="eca-empty">No employees yet — add one or use Bulk Import.</div></td></tr>}
              {filtered.map((e) => (
                <tr key={e.id}>
                  <td>{e.number}</td>
                  <td style={{ fontWeight: 600 }}>{e.lastName}</td>
                  <td>{e.firstName}</td>
                  <td>{e.office && <span className="eca-tag">{e.office}</span>}</td>
                  <td style={{ fontSize: 12, color: COLORS.gray600 }}>{e.position}</td>
                  <td style={{ whiteSpace: "nowrap" }}>
                    <button className="eca-mini-btn" onClick={() => setModal(e)}>Edit</button>
                    <button className="eca-mini-btn" onClick={() => remove(e.id)}>Remove</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      <div className="eca-help" style={{ marginTop: 10 }}>{roster.length} employees on file.</div>
    </div>
  );
}

function EmployeeModal({ initial, onSave, onClose }) {
  const [number, setNumber] = useState(initial.number || "");
  const [firstName, setFirstName] = useState(initial.firstName || "");
  const [lastName, setLastName] = useState(initial.lastName || "");
  const [office, setOffice] = useState(initial.office || "");
  const [position, setPosition] = useState(initial.position || "");

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(13,31,60,0.5)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "white", borderRadius: 12, width: "100%", maxWidth: 440, padding: 24 }}>
        <div className="eca-title" style={{ fontSize: 18, marginBottom: 16 }}>{initial.id ? "Edit Employee" : "Add Employee"}</div>
        <div className="eca-field"><label className="eca-label">Employee Number</label><input className="eca-input" value={number} onChange={(e) => setNumber(e.target.value)} /></div>
        <div className="eca-field"><label className="eca-label">First Name</label><input className="eca-input" value={firstName} onChange={(e) => setFirstName(e.target.value)} /></div>
        <div className="eca-field"><label className="eca-label">Last Name</label><input className="eca-input" value={lastName} onChange={(e) => setLastName(e.target.value)} /></div>
        <div className="eca-field">
          <label className="eca-label">Office</label>
          <select className="eca-input" value={office} onChange={(e) => setOffice(e.target.value)}>
            <option value="">—</option>
            {OFFICES.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div className="eca-field"><label className="eca-label">Position</label><input className="eca-input" value={position} onChange={(e) => setPosition(e.target.value)} placeholder="e.g. Operator, Foreman" /></div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
          <button className="eca-btn eca-btn-secondary" onClick={onClose}>Cancel</button>
          <button className="eca-btn eca-btn-primary" onClick={() => onSave({ ...initial, number, firstName, lastName, office, position })}>Save</button>
        </div>
      </div>
    </div>
  );
}

function AdminCategories({ categories, setCategories, roster, votes, refreshVotes }) {
  const [expanded, setExpanded] = useState(null);
  const [newCatName, setNewCatName] = useState("");
  const [bulkNomineeText, setBulkNomineeText] = useState({});
  const [nomineeQuery, setNomineeQuery] = useState({});
  const [officeConfirm, setOfficeConfirm] = useState(null); // { catId, office } | null
  const [positionConfirm, setPositionConfirm] = useState(null); // { catId, position } | null
  const [positionPick, setPositionPick] = useState({});
  const [editingNominee, setEditingNominee] = useState(null); // { catId, nominee } | null
  const [movingNominee, setMovingNominee] = useState(null); // { catId, nominee } | null
  const [movePick, setMovePick] = useState("");
  const [moving, setMoving] = useState(false);
  const [moveResult, setMoveResult] = useState(null); // { moved, conflicts, destName } | null

  const persist = (next) => {
    setCategories(next);
    saveCategories(next);
  };

  const addCategory = () => {
    if (!newCatName.trim()) return;
    const cat = { id: uid(), name: newCatName.trim(), winnerCount: 1, nominees: [] };
    persist([...categories, cat]);
    setNewCatName("");
    setExpanded(cat.id);
  };

  const updateCategory = (id, patch) => {
    persist(categories.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  };

  const removeCategory = (id) => {
    persist(categories.filter((c) => c.id !== id));
  };

  const moveCategory = (index, direction) => {
    const newIndex = index + direction;
    if (newIndex < 0 || newIndex >= categories.length) return;
    const next = categories.slice();
    const [moved] = next.splice(index, 1);
    next.splice(newIndex, 0, moved);
    persist(next);
  };

  const addNomineeBulk = (catId) => {
    const text = bulkNomineeText[catId] || "";
    const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
    const additions = lines.map((line) => {
      const parts = line.split(",").map((p) => p.trim());
      const [name, office] = parts;
      const { firstName, lastName } = splitName(name || "");
      return { id: uid(), firstName, lastName, office: office || "" };
    });
    persist(categories.map((c) => (c.id === catId ? { ...c, nominees: [...c.nominees, ...additions] } : c)));
    setBulkNomineeText((s) => ({ ...s, [catId]: "" }));
  };

  const addNomineeFromRoster = (catId, employee) => {
    const cat = categories.find((c) => c.id === catId);
    if (!cat) return;
    if (cat.nominees.some((n) => n.number && n.number === employee.number)) return; // already added
    const nominee = { id: uid(), number: employee.number, firstName: employee.firstName, lastName: employee.lastName, office: employee.office || "", position: employee.position || "" };
    persist(categories.map((c) => (c.id === catId ? { ...c, nominees: [...c.nominees, nominee] } : c)));
    setNomineeQuery((s) => ({ ...s, [catId]: "" }));
  };

  const addAllByOffice = (catId, office) => {
    const cat = categories.find((c) => c.id === catId);
    if (!cat) return;
    const existingNumbers = new Set(cat.nominees.filter((n) => n.number).map((n) => n.number));
    const additions = roster
      .filter((e) => e.office === office && !existingNumbers.has(e.number))
      .map((e) => ({ id: uid(), number: e.number, firstName: e.firstName, lastName: e.lastName, office: e.office, position: e.position || "" }));
    persist(categories.map((c) => (c.id === catId ? { ...c, nominees: [...c.nominees, ...additions] } : c)));
    setOfficeConfirm(null);
  };

  const addAllByPosition = (catId, position) => {
    const cat = categories.find((c) => c.id === catId);
    if (!cat) return;
    const existingNumbers = new Set(cat.nominees.filter((n) => n.number).map((n) => n.number));
    const additions = roster
      .filter((e) => e.position === position && !existingNumbers.has(e.number))
      .map((e) => ({ id: uid(), number: e.number, firstName: e.firstName, lastName: e.lastName, office: e.office || "", position: e.position }));
    persist(categories.map((c) => (c.id === catId ? { ...c, nominees: [...c.nominees, ...additions] } : c)));
    setPositionConfirm(null);
  };

  const removeNominee = (catId, nomineeId) => {
    persist(categories.map((c) => (c.id === catId ? { ...c, nominees: c.nominees.filter((n) => n.id !== nomineeId) } : c)));
  };

  // Edits a nominee's details in place. Same id, same category, so any
  // existing votes for this person are completely unaffected.
  const updateNominee = (catId, nomineeId, patch) => {
    persist(categories.map((c) =>
      c.id === catId
        ? { ...c, nominees: c.nominees.map((n) => (n.id === nomineeId ? { ...n, ...patch } : n)) }
        : c
    ));
    setEditingNominee(null);
  };

  // Moves a nominee from one category to another, keeping their same id,
  // and rewrites every existing vote record that pointed at
  // (fromCatId, nomineeId) to instead point at (toCatId, nomineeId) —
  // so their vote tally carries over instead of resetting to zero.
  // If a voter already has a separate pick in the destination category,
  // that voter's old vote for this person can't be carried over (it
  // would overwrite a real choice), so it's counted as a conflict and
  // left alone rather than silently dropped or overwritten.
  const moveNominee = async (fromCatId, toCatId, nominee) => {
    setMoving(true);
    const freshVotes = await loadAllVotes();
    let movedCount = 0;
    let conflictCount = 0;

    for (const v of freshVotes) {
      const sel = v.selections || {};
      if (sel[fromCatId] !== nominee.id) continue;
      if (sel[toCatId] !== undefined) {
        conflictCount++;
        continue;
      }
      const nextSelections = { ...sel };
      delete nextSelections[fromCatId];
      nextSelections[toCatId] = nominee.id;
      const updatedRecord = { ...v, selections: nextSelections, lastUpdated: new Date().toISOString() };
      await submitVoteRecord(v.employeeNumber, updatedRecord);
      movedCount++;
    }

    const destCat = categories.find((c) => c.id === toCatId);
    persist(categories.map((c) => {
      if (c.id === fromCatId) return { ...c, nominees: c.nominees.filter((n) => n.id !== nominee.id) };
      if (c.id === toCatId) return { ...c, nominees: [...c.nominees, nominee] };
      return c;
    }));

    setMoving(false);
    setMovingNominee(null);
    setMovePick("");
    setMoveResult({ moved: movedCount, conflicts: conflictCount, destName: destCat?.name || "" });
    await refreshVotes();
  };

  const distinctPositions = Array.from(new Set(roster.map((e) => e.position).filter(Boolean))).sort((a, b) => a.localeCompare(b));

  return (
    <div>
      <div className="eca-card" style={{ marginBottom: 18 }}>
        <div className="eca-section-title">Add a Category</div>
        <div style={{ display: "flex", gap: 8 }}>
          <input className="eca-input" placeholder="e.g. Operator / Grade Checker" value={newCatName} onChange={(e) => setNewCatName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addCategory()} />
          <button className="eca-btn eca-btn-primary eca-btn-sm" onClick={addCategory}>Add</button>
        </div>
      </div>

      {categories.length === 0 && <div className="eca-empty">No categories yet. Add your first one above.</div>}

      {categories.map((cat, catIndex) => {
        const isOpen = expanded === cat.id;
        const query = nomineeQuery[cat.id] || "";
        const existingNumbers = new Set(cat.nominees.filter((n) => n.number).map((n) => n.number));
        const searchResults = query.trim()
          ? roster
              .map((e) => ({ e, m: fuzzyMatch(query, fullName(e)) }))
              .filter((x) => x.m.match)
              .sort((a, b) => b.m.score - a.m.score || byLastName(a.e, b.e))
              .slice(0, 8)
          : [];
        return (
          <div key={cat.id} className="eca-category-block">
            <div className="eca-category-head" onClick={() => setExpanded(isOpen ? null : cat.id)}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <button
                    className="eca-mini-btn"
                    style={{ padding: "0 4px", lineHeight: 1, opacity: catIndex === 0 ? 0.3 : 1 }}
                    disabled={catIndex === 0}
                    onClick={(e) => { e.stopPropagation(); moveCategory(catIndex, -1); }}
                    title="Move up"
                  >▲</button>
                  <button
                    className="eca-mini-btn"
                    style={{ padding: "0 4px", lineHeight: 1, opacity: catIndex === categories.length - 1 ? 0.3 : 1 }}
                    disabled={catIndex === categories.length - 1}
                    onClick={(e) => { e.stopPropagation(); moveCategory(catIndex, 1); }}
                    title="Move down"
                  >▼</button>
                </div>
                <div>
                  <div className="eca-category-name">{cat.name}</div>
                  <div className="eca-category-meta">{cat.nominees.length} nominees · {cat.winnerCount} winner{cat.winnerCount === 1 ? "" : "s"}</div>
                </div>
              </div>
              <button className="eca-mini-btn" onClick={(e) => { e.stopPropagation(); removeCategory(cat.id); }}>Remove Category</button>
            </div>
            {isOpen && (
              <div style={{ padding: "6px 18px 18px" }}>
                <div style={{ display: "flex", gap: 16, marginBottom: 16, flexWrap: "wrap" }}>
                  <div className="eca-field" style={{ marginBottom: 0, flex: "1 1 220px" }}>
                    <label className="eca-label">Category Name</label>
                    <input className="eca-input" value={cat.name} onChange={(e) => updateCategory(cat.id, { name: e.target.value })} />
                  </div>
                  <div className="eca-field" style={{ marginBottom: 0, width: 140 }}>
                    <label className="eca-label"># of Winners</label>
                    <input type="number" min="1" className="eca-input" value={cat.winnerCount} onChange={(e) => updateCategory(cat.id, { winnerCount: +e.target.value || 1 })} />
                  </div>
                </div>

                <div className="eca-section-title" style={{ fontSize: 14 }}>Nominees</div>
                <div className="eca-table-wrap" style={{ marginBottom: 12 }}>
                  <table className="eca-table">
                    <thead><tr><th>Last</th><th>First</th><th>Office</th><th>Position</th><th></th></tr></thead>
                    <tbody>
                      {cat.nominees.length === 0 && <tr><td colSpan={5}><div className="eca-empty">No nominees added yet.</div></td></tr>}
                      {cat.nominees.slice().sort(byLastName).map((n) => (
                        <tr key={n.id}>
                          <td style={{ fontWeight: 600 }}>{n.lastName}</td>
                          <td>{n.firstName}</td>
                          <td>{n.office && <span className="eca-tag">{n.office}</span>}</td>
                          <td style={{ fontSize: 12, color: COLORS.gray600 }}>{n.position}</td>
                          <td style={{ whiteSpace: "nowrap" }}>
                            <button className="eca-mini-btn" onClick={() => setEditingNominee({ catId: cat.id, nominee: n })}>Edit</button>
                            <button className="eca-mini-btn" onClick={() => { setMovingNominee({ catId: cat.id, nominee: n }); setMovePick(""); }}>Move</button>
                            <button className="eca-mini-btn" onClick={() => removeNominee(cat.id, n.id)}>Remove</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <label className="eca-label">Add From Employee Roster</label>
                <div style={{ position: "relative", marginBottom: 14 }}>
                  <input
                    className="eca-input"
                    placeholder="Start typing a name…"
                    value={query}
                    onChange={(e) => setNomineeQuery((s) => ({ ...s, [cat.id]: e.target.value }))}
                  />
                  {searchResults.length > 0 && (
                    <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: "white", border: `1px solid ${COLORS.gray200}`, borderRadius: 8, marginTop: 4, zIndex: 10, boxShadow: "0 6px 20px rgba(13,31,60,0.1)", maxHeight: 260, overflowY: "auto" }}>
                      {searchResults.map(({ e }) => {
                        const already = existingNumbers.has(e.number);
                        return (
                          <div
                            key={e.number}
                            onClick={() => !already && addNomineeFromRoster(cat.id, e)}
                            style={{ padding: "9px 14px", cursor: already ? "default" : "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `1px solid ${COLORS.gray100}`, opacity: already ? 0.5 : 1 }}
                          >
                            <span style={{ fontSize: 13, color: COLORS.navy }}>
                              {fullName(e)} <span style={{ color: COLORS.gray400 }}>#{e.number}{e.position ? ` · ${e.position}` : ""}</span>
                            </span>
                            {already ? <span className="eca-skip-pill">Added</span> : (e.office && <span className="eca-tag">{e.office}</span>)}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>

                <label className="eca-label">Add Everyone From an Office</label>
                <div className="eca-help" style={{ marginBottom: 8 }}>
                  Loads every roster employee from that office into this category as a nominee. Skips anyone already added.
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
                  {OFFICES.map((office) => {
                    const count = roster.filter((e) => e.office === office && !existingNumbers.has(e.number)).length;
                    return (
                      <button
                        key={office}
                        className="eca-btn eca-btn-secondary eca-btn-sm"
                        disabled={count === 0}
                        onClick={() => setOfficeConfirm({ catId: cat.id, office })}
                      >
                        + Add all {office} ({count})
                      </button>
                    );
                  })}
                </div>

                {officeConfirm && officeConfirm.catId === cat.id && (
                  <div className="eca-error" style={{ marginBottom: 14 }}>
                    Add all {roster.filter((e) => e.office === officeConfirm.office && !existingNumbers.has(e.number)).length} {officeConfirm.office} employees to "{cat.name}"?
                    <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                      <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={() => addAllByOffice(cat.id, officeConfirm.office)}>Yes, add them all</button>
                      <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setOfficeConfirm(null)}>Cancel</button>
                    </div>
                  </div>
                )}

                <label className="eca-label">Add Everyone With a Position</label>
                <div className="eca-help" style={{ marginBottom: 8 }}>
                  Loads every roster employee whose job title matches the one you pick (e.g. all Operators, all Foremen). Pick more than one position, one at a time, to build a combined category like Operator/Grade Checker. Skips anyone already added.
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16, alignItems: "center" }}>
                  <select
                    className="eca-input"
                    style={{ maxWidth: 260 }}
                    value={positionPick[cat.id] || ""}
                    onChange={(e) => setPositionPick((s) => ({ ...s, [cat.id]: e.target.value }))}
                  >
                    <option value="">Select a position…</option>
                    {distinctPositions.map((p) => {
                      const count = roster.filter((e) => e.position === p && !existingNumbers.has(e.number)).length;
                      return <option key={p} value={p}>{p} ({count} available)</option>;
                    })}
                  </select>
                  <button
                    className="eca-btn eca-btn-secondary eca-btn-sm"
                    disabled={!positionPick[cat.id]}
                    onClick={() => setPositionConfirm({ catId: cat.id, position: positionPick[cat.id] })}
                  >
                    + Add Everyone in This Position
                  </button>
                </div>

                {positionConfirm && positionConfirm.catId === cat.id && (
                  <div className="eca-error" style={{ marginBottom: 14 }}>
                    Add all {roster.filter((e) => e.position === positionConfirm.position && !existingNumbers.has(e.number)).length} employees with position "{positionConfirm.position}" to "{cat.name}"?
                    <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                      <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={() => addAllByPosition(cat.id, positionConfirm.position)}>Yes, add them all</button>
                      <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setPositionConfirm(null)}>Cancel</button>
                    </div>
                  </div>
                )}

                <label className="eca-label">Bulk Add Nominees Manually (one per line: name, office)</label>
                <div className="eca-help" style={{ marginBottom: 6 }}>
                  Use this for anyone not on the roster. Names added this way aren't linked to a roster record.
                </div>
                <textarea
                  className="eca-textarea"
                  style={{ minHeight: 70 }}
                  value={bulkNomineeText[cat.id] || ""}
                  onChange={(e) => setBulkNomineeText((s) => ({ ...s, [cat.id]: e.target.value }))}
                  placeholder="John Smith, Tucson"
                />
                <button className="eca-btn eca-btn-secondary eca-btn-sm" style={{ marginTop: 8 }} onClick={() => addNomineeBulk(cat.id)}>Add Nominees</button>
              </div>
            )}
          </div>
        );
      })}

      {editingNominee && (
        <EditNomineeModal
          initial={editingNominee.nominee}
          onSave={(patch) => updateNominee(editingNominee.catId, editingNominee.nominee.id, patch)}
          onClose={() => setEditingNominee(null)}
        />
      )}

      {movingNominee && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(13,31,60,0.5)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
          <div style={{ background: "white", borderRadius: 12, width: "100%", maxWidth: 440, padding: 24 }}>
            <div className="eca-title" style={{ fontSize: 18, marginBottom: 6 }}>Move Nominee</div>
            <div className="eca-help" style={{ marginBottom: 16 }}>
              Moving <strong>{fullName(movingNominee.nominee)}</strong> out of "{categories.find((c) => c.id === movingNominee.catId)?.name}". Any votes they've already received in that category will move with them, as long as the voter hasn't also voted for someone else in the destination category.
            </div>
            <div className="eca-field">
              <label className="eca-label">Move to Category</label>
              <select className="eca-input" value={movePick} onChange={(e) => setMovePick(e.target.value)}>
                <option value="">— Select category —</option>
                {categories.filter((c) => c.id !== movingNominee.catId).map((c) => (
                  <option key={c.id} value={c.id}>{c.name}</option>
                ))}
              </select>
            </div>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
              <button className="eca-btn eca-btn-secondary" onClick={() => { setMovingNominee(null); setMovePick(""); }}>Cancel</button>
              <button
                className="eca-btn eca-btn-primary"
                disabled={!movePick || moving}
                onClick={() => moveNominee(movingNominee.catId, movePick, movingNominee.nominee)}
              >
                {moving ? "Moving…" : "Move & Carry Over Votes"}
              </button>
            </div>
          </div>
        </div>
      )}

      {moveResult && (
        <div style={{ position: "fixed", bottom: 24, right: 24, zIndex: 250, maxWidth: 340 }}>
          <div className="eca-success" style={{ marginBottom: 0, boxShadow: "0 6px 20px rgba(13,31,60,0.15)" }}>
            Moved to {moveResult.destName}. {moveResult.moved} vote{moveResult.moved === 1 ? "" : "s"} carried over.
            {moveResult.conflicts > 0 && ` ${moveResult.conflicts} vote${moveResult.conflicts === 1 ? "" : "s"} couldn't move because that voter already picked someone else there.`}
            <div style={{ marginTop: 8 }}>
              <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setMoveResult(null)}>Dismiss</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function EditNomineeModal({ initial, onSave, onClose }) {
  const [firstName, setFirstName] = useState(initial.firstName || "");
  const [lastName, setLastName] = useState(initial.lastName || "");
  const [office, setOffice] = useState(initial.office || "");
  const [position, setPosition] = useState(initial.position || "");

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(13,31,60,0.5)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "white", borderRadius: 12, width: "100%", maxWidth: 440, padding: 24 }}>
        <div className="eca-title" style={{ fontSize: 18, marginBottom: 16 }}>Edit Nominee</div>
        <div className="eca-help" style={{ marginBottom: 14 }}>
          Changing how their name displays doesn't affect their votes — same person, same tally, just corrected spelling or the name they go by.
        </div>
        <div className="eca-field"><label className="eca-label">First Name</label><input className="eca-input" value={firstName} onChange={(e) => setFirstName(e.target.value)} /></div>
        <div className="eca-field"><label className="eca-label">Last Name</label><input className="eca-input" value={lastName} onChange={(e) => setLastName(e.target.value)} /></div>
        <div className="eca-field">
          <label className="eca-label">Office</label>
          <select className="eca-input" value={office} onChange={(e) => setOffice(e.target.value)}>
            <option value="">—</option>
            {OFFICES.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div className="eca-field"><label className="eca-label">Position</label><input className="eca-input" value={position} onChange={(e) => setPosition(e.target.value)} /></div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 6 }}>
          <button className="eca-btn eca-btn-secondary" onClick={onClose}>Cancel</button>
          <button className="eca-btn eca-btn-primary" onClick={() => onSave({ firstName, lastName, office, position })}>Save</button>
        </div>
      </div>
    </div>
  );
}

function AdminVotes({ roster, categories, votes, refreshVotes }) {
  const [search, setSearch] = useState("");
  const [selectedNumber, setSelectedNumber] = useState(null);
  const [draft, setDraft] = useState({});
  const [loadingBallot, setLoadingBallot] = useState(false);
  const [existingRecord, setExistingRecord] = useState(null);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState("");
  const [confirmClear, setConfirmClear] = useState(false);

  // Nominee-centric editing: pick a category + nominee, see who voted for
  // them, remove individual votes, or assign a new vote to them.
  const [nomCatId, setNomCatId] = useState("");
  const [nomNomineeId, setNomNomineeId] = useState("");
  const [nomMsg, setNomMsg] = useState("");
  const [addVoteSearch, setAddVoteSearch] = useState("");
  const [addingVoteBusy, setAddingVoteBusy] = useState(false);
  const [removingVoterNumber, setRemovingVoterNumber] = useState(null);
  const [addVotePending, setAddVotePending] = useState(null); // { employee, currentChoiceName } | null

  const votedNumbers = new Set(votes.map((v) => String(v.employeeNumber)));

  const results = search.trim()
    ? roster
        .map((e) => ({ e, m: fuzzyMatch(search, `${fullName(e)} ${e.number}`) }))
        .filter((x) => x.m.match)
        .sort((a, b) => b.m.score - a.m.score || byLastName(a.e, b.e))
        .slice(0, 10)
    : [];

  const openVoter = async (employee) => {
    setMsg("");
    setSelectedNumber(employee.number);
    setLoadingBallot(true);
    const record = await getVoteRecord(employee.number);
    setExistingRecord(record);
    setDraft(record?.selections ? { ...record.selections } : {});
    setLoadingBallot(false);
  };

  const setChoice = (catId, nomineeId) => {
    setDraft((d) => {
      const next = { ...d };
      if (!nomineeId) delete next[catId];
      else next[catId] = nomineeId;
      return next;
    });
  };

  const selectedEmployee = roster.find((e) => e.number === selectedNumber);

  const saveBallot = async () => {
    if (!selectedEmployee) return;
    setSaving(true);
    const record = {
      employeeNumber: selectedEmployee.number,
      voterFirstName: selectedEmployee.firstName,
      voterLastName: selectedEmployee.lastName,
      office: selectedEmployee.office || "",
      timestamp: existingRecord?.timestamp || new Date().toISOString(),
      lastUpdated: new Date().toISOString(),
      selections: draft,
    };
    await submitVoteRecord(selectedEmployee.number, record);
    setSaving(false);
    setMsg(`Saved ${fullName(selectedEmployee)}'s ballot.`);
    await refreshVotes();
    setExistingRecord(record);
  };

  const clearBallot = async () => {
    if (!selectedEmployee) return;
    setSaving(true);
    await safeSet(VOTE_PREFIX + selectedEmployee.number, "", true);
    setSaving(false);
    setConfirmClear(false);
    setDraft({});
    setExistingRecord(null);
    setMsg(`Cleared ${fullName(selectedEmployee)}'s ballot — they now show as not voted.`);
    await refreshVotes();
  };

  const nomCategory = categories.find((c) => c.id === nomCatId);
  const nomNominee = nomCategory?.nominees.find((n) => n.id === nomNomineeId);

  // Everyone currently credited with voting for this nominee in this category.
  const votersForNominee = nomCatId && nomNomineeId
    ? votes.filter((v) => v.selections && v.selections[nomCatId] === nomNomineeId)
    : [];

  const removeVoteFor = async (voteRecord) => {
    setRemovingVoterNumber(voteRecord.employeeNumber);
    const fresh = await getVoteRecord(voteRecord.employeeNumber);
    if (fresh) {
      const nextSelections = { ...fresh.selections };
      delete nextSelections[nomCatId];
      await submitVoteRecord(voteRecord.employeeNumber, { ...fresh, selections: nextSelections, lastUpdated: new Date().toISOString() });
    }
    setRemovingVoterNumber(null);
    setNomMsg(`Removed ${voteRecord.voterFirstName} ${voteRecord.voterLastName}'s vote for ${fullName(nomNominee)}.`);
    await refreshVotes();
  };

  const votersForNomineeNumbers = new Set(votersForNominee.map((v) => String(v.employeeNumber)));
  const addVoteResults = addVoteSearch.trim() && nomCatId
    ? roster
        .filter((e) => !votersForNomineeNumbers.has(String(e.number)))
        .map((e) => ({ e, m: fuzzyMatch(addVoteSearch, `${fullName(e)} ${e.number}`) }))
        .filter((x) => x.m.match)
        .sort((a, b) => b.m.score - a.m.score || byLastName(a.e, b.e))
        .slice(0, 8)
    : [];

  const startAddVote = async (employee) => {
    const fresh = await getVoteRecord(employee.number);
    const currentNomineeId = fresh?.selections?.[nomCatId];
    const currentChoiceName = currentNomineeId ? fullName(nomCategory.nominees.find((n) => n.id === currentNomineeId)) : null;
    setAddVotePending({ employee, currentChoiceName });
  };

  const assignVote = async (employee) => {
    setAddingVoteBusy(true);
    const fresh = await getVoteRecord(employee.number);
    const record = fresh || {
      employeeNumber: employee.number,
      voterFirstName: employee.firstName,
      voterLastName: employee.lastName,
      office: employee.office || "",
      timestamp: new Date().toISOString(),
      selections: {},
    };
    const nextSelections = { ...record.selections, [nomCatId]: nomNomineeId };
    await submitVoteRecord(employee.number, { ...record, selections: nextSelections, lastUpdated: new Date().toISOString() });
    setAddingVoteBusy(false);
    setAddVotePending(null);
    setAddVoteSearch("");
    setNomMsg(`Added ${fullName(employee)}'s vote for ${fullName(nomNominee)}.`);
    await refreshVotes();
  };

  return (
    <div>
      <div className="eca-card" style={{ marginBottom: 18 }}>
        <div className="eca-section-title">Find a Voter</div>
        <div className="eca-help" style={{ marginBottom: 10 }}>
          Search by name or employee number to view or manually correct someone's ballot — useful for fixing a miscast vote or entering a paper ballot on someone's behalf.
        </div>
        <input
          className="eca-input"
          placeholder="Search by name or employee number…"
          value={search}
          onChange={(e) => { setSearch(e.target.value); setMsg(""); }}
        />
        {results.length > 0 && (
          <div style={{ marginTop: 10, border: `1px solid ${COLORS.gray200}`, borderRadius: 8, maxHeight: 260, overflowY: "auto" }}>
            {results.map(({ e }) => (
              <div
                key={e.number}
                onClick={() => openVoter(e)}
                style={{ padding: "9px 14px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `1px solid ${COLORS.gray100}` }}
              >
                <span style={{ fontSize: 13, color: COLORS.navy }}>{fullName(e)} <span style={{ color: COLORS.gray400 }}>#{e.number}</span></span>
                {votedNumbers.has(String(e.number)) ? <span className="eca-chosen-pill">Voted</span> : <span className="eca-skip-pill">Not voted</span>}
              </div>
            ))}
          </div>
        )}
      </div>

      {msg && <div className="eca-success">{msg}</div>}

      {selectedEmployee && (
        <div className="eca-card">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
            <div>
              <div className="eca-section-title" style={{ marginBottom: 2 }}>{fullName(selectedEmployee)}</div>
              <div className="eca-help" style={{ marginBottom: 0 }}>
                #{selectedEmployee.number} · {selectedEmployee.office || "No office"}
                {existingRecord ? " · has a submitted ballot" : " · hasn't voted yet"}
              </div>
            </div>
            {existingRecord && (
              !confirmClear ? (
                <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmClear(true)}>Clear Their Ballot</button>
              ) : (
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 12, color: COLORS.red, marginBottom: 6 }}>Reset to not-voted?</div>
                  <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={clearBallot} style={{ marginRight: 6 }}>Yes, clear</button>
                  <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmClear(false)}>Cancel</button>
                </div>
              )
            )}
          </div>

          <div className="eca-divider" />

          {loadingBallot ? (
            <div className="eca-empty">Loading ballot…</div>
          ) : (
            <>
              {categories.length === 0 && <div className="eca-empty">No categories set up yet.</div>}
              {categories.map((cat) => (
                <div key={cat.id} className="eca-field">
                  <label className="eca-label">{cat.name}</label>
                  <select className="eca-input" value={draft[cat.id] || ""} onChange={(e) => setChoice(cat.id, e.target.value)}>
                    <option value="">— No selection —</option>
                    {cat.nominees.slice().sort(byLastName).map((n) => (
                      <option key={n.id} value={n.id}>{fullName(n)}{n.office ? ` — ${n.office}` : ""}</option>
                    ))}
                  </select>
                </div>
              ))}
              <button className="eca-btn eca-btn-primary" style={{ marginTop: 8 }} onClick={saveBallot} disabled={saving}>
                {saving ? "Saving…" : "Save Ballot"}
              </button>
            </>
          )}
        </div>
      )}

      <div className="eca-divider" />

      <div className="eca-card">
        <div className="eca-section-title">Edit Votes by Nominee</div>
        <div className="eca-help" style={{ marginBottom: 14 }}>
          Pick a category and a nominee to see everyone currently credited with voting for them. Remove individual votes, or add a vote on someone's behalf.
        </div>

        <div style={{ display: "flex", gap: 16, marginBottom: 16, flexWrap: "wrap" }}>
          <div className="eca-field" style={{ marginBottom: 0, flex: "1 1 220px" }}>
            <label className="eca-label">Category</label>
            <select
              className="eca-input"
              value={nomCatId}
              onChange={(e) => { setNomCatId(e.target.value); setNomNomineeId(""); setNomMsg(""); setAddVotePending(null); }}
            >
              <option value="">— Select category —</option>
              {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </div>
          <div className="eca-field" style={{ marginBottom: 0, flex: "1 1 220px" }}>
            <label className="eca-label">Nominee</label>
            <select
              className="eca-input"
              value={nomNomineeId}
              onChange={(e) => { setNomNomineeId(e.target.value); setNomMsg(""); setAddVotePending(null); }}
              disabled={!nomCatId}
            >
              <option value="">— Select nominee —</option>
              {nomCategory?.nominees.slice().sort(byLastName).map((n) => (
                <option key={n.id} value={n.id}>{fullName(n)}</option>
              ))}
            </select>
          </div>
        </div>

        {nomMsg && <div className="eca-success">{nomMsg}</div>}

        {nomCatId && nomNomineeId && (
          <>
            <div className="eca-section-title" style={{ fontSize: 14 }}>
              Current Votes ({votersForNominee.length})
            </div>
            <div className="eca-table-wrap" style={{ marginBottom: 18 }}>
              <table className="eca-table">
                <thead><tr><th>Voter</th><th>Employee #</th><th></th></tr></thead>
                <tbody>
                  {votersForNominee.length === 0 && <tr><td colSpan={3}><div className="eca-empty">No one has voted for {fullName(nomNominee)} in this category yet.</div></td></tr>}
                  {votersForNominee.map((v) => (
                    <tr key={v.employeeNumber}>
                      <td className="td-name" style={{ fontWeight: 600 }}>{v.voterFirstName} {v.voterLastName}</td>
                      <td>{v.employeeNumber}</td>
                      <td>
                        <button className="eca-mini-btn" disabled={removingVoterNumber === v.employeeNumber} onClick={() => removeVoteFor(v)}>
                          {removingVoterNumber === v.employeeNumber ? "Removing…" : "Remove Vote"}
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <label className="eca-label">Add a Vote for {fullName(nomNominee)}</label>
            <div className="eca-help" style={{ marginBottom: 8 }}>
              Search for the employee whose vote you want to assign here.
            </div>
            <div style={{ position: "relative", marginBottom: 14 }}>
              <input
                className="eca-input"
                placeholder="Start typing a name or employee number…"
                value={addVoteSearch}
                onChange={(e) => { setAddVoteSearch(e.target.value); setAddVotePending(null); }}
              />
              {addVoteResults.length > 0 && !addVotePending && (
                <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: "white", border: `1px solid ${COLORS.gray200}`, borderRadius: 8, marginTop: 4, zIndex: 10, boxShadow: "0 6px 20px rgba(13,31,60,0.1)", maxHeight: 260, overflowY: "auto" }}>
                  {addVoteResults.map(({ e }) => (
                    <div
                      key={e.number}
                      onClick={() => startAddVote(e)}
                      style={{ padding: "9px 14px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: `1px solid ${COLORS.gray100}` }}
                    >
                      <span style={{ fontSize: 13, color: COLORS.navy }}>{fullName(e)} <span style={{ color: COLORS.gray400 }}>#{e.number}</span></span>
                      {e.office && <span className="eca-tag">{e.office}</span>}
                    </div>
                  ))}
                </div>
              )}
            </div>

            {addVotePending && (
              <div className={addVotePending.currentChoiceName ? "eca-error" : "eca-success"} style={{ marginBottom: 12 }}>
                {addVotePending.currentChoiceName
                  ? `${fullName(addVotePending.employee)} already voted for ${addVotePending.currentChoiceName} in this category. Assigning this vote will change their pick to ${fullName(nomNominee)}.`
                  : `${fullName(addVotePending.employee)} hasn't voted in this category yet. This will add their vote for ${fullName(nomNominee)}.`}
                <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                  <button className="eca-btn eca-btn-primary eca-btn-sm" disabled={addingVoteBusy} onClick={() => assignVote(addVotePending.employee)}>
                    {addingVoteBusy ? "Saving…" : "Confirm"}
                  </button>
                  <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setAddVotePending(null)}>Cancel</button>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function AdminExport({ roster, categories, votes }) {
  const exportResults = () => {
    const rows = [["Category", "Nominee Last Name", "Nominee First Name", "Office", "Position", "Votes Received"]];
    categories.forEach((cat) => {
      const tally = {};
      votes.forEach((v) => {
        const choice = v.selections && v.selections[cat.id];
        if (choice) tally[choice] = (tally[choice] || 0) + 1;
      });
      const sortedNominees = cat.nominees.slice().sort(byLastName);
      sortedNominees.forEach((n) => {
        rows.push([cat.name, n.lastName, n.firstName, n.office || "", n.position || "", tally[n.id] || 0]);
      });
    });
    downloadCSV(`KEG_Award_Results_${new Date().toISOString().slice(0, 10)}.csv`, rows);
  };

  const exportParticipation = () => {
    const rows = [["Last Name", "First Name", "Employee Number", "Office", "Submitted At"]];
    const sorted = votes.slice().sort((a, b) => {
      const la = (a.voterLastName || "").toLowerCase();
      const lb = (b.voterLastName || "").toLowerCase();
      return la < lb ? -1 : la > lb ? 1 : 0;
    });
    sorted.forEach((v) => {
      rows.push([v.voterLastName, v.voterFirstName, v.employeeNumber, v.office || "", v.timestamp]);
    });
    downloadCSV(`KEG_Voter_Participation_${new Date().toISOString().slice(0, 10)}.csv`, rows);
  };

  const exportNonVoters = () => {
    const votedNumbers = new Set(votes.map((v) => String(v.employeeNumber)));
    const rows = [["Last Name", "First Name", "Employee Number", "Office"]];
    roster
      .filter((e) => !votedNumbers.has(String(e.number)))
      .slice()
      .sort(byLastName)
      .forEach((e) => rows.push([e.lastName, e.firstName, e.number, e.office || ""]));
    downloadCSV(`KEG_Have_Not_Voted_${new Date().toISOString().slice(0, 10)}.csv`, rows);
  };

  return (
    <div className="eca-card">
      <div className="eca-section-title">Export</div>
      <div className="eca-help" style={{ marginBottom: 18 }}>
        Results are grouped by category with nominees sorted alphabetically by last name, plus their vote count so you can determine winners (including any office split you want to apply manually).
      </div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        <button className="eca-btn eca-btn-primary eca-btn-sm" onClick={exportResults}>Export Results by Category (CSV)</button>
        <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={exportParticipation}>Export Voter Participation (CSV)</button>
        <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={exportNonVoters}>Export Who Hasn't Voted (CSV)</button>
      </div>
    </div>
  );
}

function AdminSettings({ storedPasscode, onSetPasscode, roster, categories, setRoster, setCategories, votes, onVotesReset }) {
  const [newCode, setNewCode] = useState("");
  const [confirmNewCode, setNewCodeConfirm] = useState("");
  const [msg, setMsg] = useState("");
  const [confirmReset, setConfirmReset] = useState(false);
  const [confirmWipe, setConfirmWipe] = useState(false);
  const [backingUp, setBackingUp] = useState(false);
  const [restoreError, setRestoreError] = useState("");
  const [restorePreview, setRestorePreview] = useState(null);
  const [restoring, setRestoring] = useState(false);

  const downloadBackup = async () => {
    setBackingUp(true);
    const freshVotes = await loadAllVotes();
    downloadJSON(`KEG_ECA_Backup_${new Date().toISOString().slice(0, 10)}.json`, {
      exportedAt: new Date().toISOString(),
      roster,
      categories,
      votes: freshVotes,
    });
    setBackingUp(false);
  };

  const handleBackupFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setRestoreError("");
    setRestorePreview(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result);
        if (!Array.isArray(parsed.roster) || !Array.isArray(parsed.categories) || !Array.isArray(parsed.votes)) {
          setRestoreError("That file doesn't look like a KE&G ECA backup — missing roster, categories, or votes.");
          return;
        }
        setRestorePreview(parsed);
      } catch (err) {
        setRestoreError("Couldn't read that file as JSON.");
      }
    };
    reader.readAsText(file);
  };

  const runRestore = async () => {
    if (!restorePreview) return;
    setRestoring(true);
    await saveRoster(restorePreview.roster);
    await saveCategories(restorePreview.categories);
    for (const v of restorePreview.votes) {
      await submitVoteRecord(v.employeeNumber, v);
    }
    setRoster(restorePreview.roster);
    setCategories(restorePreview.categories);
    await onVotesReset();
    setRestoring(false);
    setRestorePreview(null);
    setMsg(`Restored from backup: ${restorePreview.roster.length} employees, ${restorePreview.categories.length} categories, ${restorePreview.votes.length} votes.`);
  };

  const changePasscode = async () => {
    if (newCode.length < 4) { setMsg("Passcode must be at least 4 characters."); return; }
    if (newCode !== confirmNewCode) { setMsg("Passcodes don't match."); return; }
    await onSetPasscode(newCode);
    setMsg("Passcode updated.");
    setNewCode(""); setNewCodeConfirm("");
  };

  const resetForNewYear = async () => {
    for (const v of votes) {
      await safeSet(VOTE_PREFIX + v.employeeNumber, "", true);
    }
    onVotesReset();
    setConfirmReset(false);
    setMsg("All votes cleared. Employee roster and categories were kept.");
  };

  const wipeEverything = async () => {
    await saveRoster([]);
    await saveCategories([]);
    for (const v of votes) {
      await safeSet(VOTE_PREFIX + v.employeeNumber, "", true);
    }
    setRoster([]);
    setCategories([]);
    onVotesReset();
    setConfirmWipe(false);
    setMsg("Everything was cleared — roster, categories, and votes.");
  };

  return (
    <div>
      {msg && <div className="eca-success">{msg}</div>}

      <div className="eca-card" style={{ marginBottom: 18 }}>
        <div className="eca-section-title">Backup &amp; Restore</div>
        <div className="eca-help" style={{ marginBottom: 12 }}>
          Download everything — every employee, every category and nominee, every vote — as one file you control. Redeploying updated code from GitHub never touches this data (code and data are separate systems), but it's worth having your own copy, especially once real votes are coming in.
        </div>
        <button className="eca-btn eca-btn-primary eca-btn-sm" onClick={downloadBackup} disabled={backingUp}>
          {backingUp ? "Preparing…" : "Download Full Backup"}
        </button>

        <div className="eca-divider" />

        <div className="eca-section-title" style={{ fontSize: 14 }}>Restore From a Backup</div>
        <div className="eca-help" style={{ marginBottom: 10 }}>
          Loads a previously downloaded backup file back into the app. This replaces the current roster and categories, and merges votes back in — use this to undo a mistake or recover from a problem.
        </div>
        {restoreError && <div className="eca-error">{restoreError}</div>}
        <input type="file" accept="application/json" onChange={handleBackupFile} style={{ fontSize: 13, marginBottom: 12 }} />
        {restorePreview && (
          <div>
            <div className="eca-error">
              This file has {restorePreview.roster.length} employees, {restorePreview.categories.length} categories, and {restorePreview.votes.length} votes
              {restorePreview.exportedAt ? ` (backed up ${new Date(restorePreview.exportedAt).toLocaleString()})` : ""}.
              Restoring will overwrite the current roster and categories, and write these votes back in. This can't be undone.
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={runRestore} disabled={restoring}>{restoring ? "Restoring…" : "Yes, restore this backup"}</button>
              <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setRestorePreview(null)}>Cancel</button>
            </div>
          </div>
        )}
      </div>

      <div className="eca-card" style={{ marginBottom: 18 }}>
        <div className="eca-section-title">Change Admin Passcode</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div className="eca-field" style={{ marginBottom: 0 }}>
            <label className="eca-label">New Passcode</label>
            <input type="password" className="eca-input" value={newCode} onChange={(e) => setNewCode(e.target.value)} />
          </div>
          <div className="eca-field" style={{ marginBottom: 0 }}>
            <label className="eca-label">Confirm</label>
            <input type="password" className="eca-input" value={confirmNewCode} onChange={(e) => setNewCodeConfirm(e.target.value)} />
          </div>
          <button className="eca-btn eca-btn-primary eca-btn-sm" onClick={changePasscode}>Update</button>
        </div>
      </div>

      <div className="eca-card" style={{ marginBottom: 18 }}>
        <div className="eca-section-title">Reset for Next Year</div>
        <div className="eca-help" style={{ marginBottom: 12 }}>
          Clears all submitted ballots so employees can vote again next year. Keeps your employee roster and categories/nominees in place.
        </div>
        {!confirmReset ? (
          <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmReset(true)}>Clear All Votes</button>
        ) : (
          <div>
            <div className="eca-error">This will permanently erase {votes.length} submitted ballots. This can't be undone.</div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={resetForNewYear}>Yes, clear all votes</button>
              <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmReset(false)}>Cancel</button>
            </div>
          </div>
        )}
      </div>

      <div className="eca-card">
        <div className="eca-section-title">Full Reset</div>
        <div className="eca-help" style={{ marginBottom: 12 }}>
          Clears everything — roster, categories, nominees, and votes. Use this only if you're starting completely over.
        </div>
        {!confirmWipe ? (
          <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmWipe(true)}>Wipe Everything</button>
        ) : (
          <div>
            <div className="eca-error">This deletes the entire roster, all categories and nominees, and all votes. This can't be undone.</div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="eca-btn eca-btn-danger eca-btn-sm" onClick={wipeEverything}>Yes, wipe everything</button>
              <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={() => setConfirmWipe(false)}>Cancel</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function AdminPanel({ roster, setRoster, categories, setCategories, storedPasscode, onSetPasscode }) {
  const [tab, setTab] = useState("overview");
  const [votes, setVotes] = useState([]);
  const [votesLoading, setVotesLoading] = useState(true);

  const refreshVotes = useCallback(async () => {
    setVotesLoading(true);
    const v = await loadAllVotes();
    setVotes(v);
    setVotesLoading(false);
  }, []);

  useEffect(() => { refreshVotes(); }, [refreshVotes]);

  const tabs = [
    { id: "overview", label: "Overview" },
    { id: "import", label: "Import" },
    { id: "employees", label: "Employees" },
    { id: "categories", label: "Categories" },
    { id: "votes", label: "Votes" },
    { id: "export", label: "Export" },
    { id: "settings", label: "Settings" },
  ];

  return (
    <div className="eca-shell-wide">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div className="eca-title" style={{ marginBottom: 0 }}>Admin</div>
        <button className="eca-btn eca-btn-secondary eca-btn-sm" onClick={refreshVotes} disabled={votesLoading}>
          {votesLoading ? "Refreshing…" : "Refresh Vote Data"}
        </button>
      </div>
      <div className="eca-tabs">
        {tabs.map((t) => (
          <div key={t.id} className={`eca-tab${tab === t.id ? " active" : ""}`} onClick={() => setTab(t.id)}>{t.label}</div>
        ))}
      </div>

      {tab === "overview" && <AdminOverview roster={roster} categories={categories} votes={votes} />}
      {tab === "import" && (
        <AdminImport
          roster={roster}
          setRoster={setRoster}
          categories={categories}
          setCategories={setCategories}
          votes={votes}
          onVotesReset={refreshVotes}
        />
      )}
      {tab === "employees" && <AdminEmployees roster={roster} setRoster={setRoster} />}
      {tab === "categories" && <AdminCategories categories={categories} setCategories={setCategories} roster={roster} votes={votes} refreshVotes={refreshVotes} />}
      {tab === "votes" && <AdminVotes roster={roster} categories={categories} votes={votes} refreshVotes={refreshVotes} />}
      {tab === "export" && <AdminExport roster={roster} categories={categories} votes={votes} />}
      {tab === "settings" && (
        <AdminSettings
          storedPasscode={storedPasscode}
          onSetPasscode={onSetPasscode}
          roster={roster}
          categories={categories}
          setRoster={setRoster}
          setCategories={setCategories}
          votes={votes}
          onVotesReset={refreshVotes}
        />
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════
// ROOT APP
// ══════════════════════════════════════════════════════════════════════

function App() {
  const [loading, setLoading] = useState(true);
  const [roster, setRoster] = useState([]);
  const [categories, setCategories] = useState([]);
  const [passcode, setPasscodeState] = useState(null);

  const [mode, setMode] = useState("voter"); // 'voter' | 'admin'
  const [adminAuthed, setAdminAuthed] = useState(false);

  const [voterStep, setVoterStep] = useState("login"); // 'login' | 'voting' | 'done'
  const [verifiedVoter, setVerifiedVoter] = useState(null);
  const [existingRecord, setExistingRecord] = useState(null);

  useEffect(() => {
    (async () => {
      const [r, c, p] = await Promise.all([loadRoster(), loadCategories(), loadPasscode()]);
      setRoster(r);
      setCategories(c);
      setPasscodeState(p);
      setLoading(false);
    })();
  }, []);

  const handleSetPasscode = async (code) => {
    await savePasscode(code);
    setPasscodeState(code);
  };

  const switchMode = () => {
    if (mode === "voter") {
      setMode("admin");
    } else {
      setMode("voter");
      setAdminAuthed(false);
    }
  };

  if (loading) {
    return (
      <>
        <style>{css}</style>
        <div className="eca-root">
          <Header mode={mode} onSwitch={switchMode} />
          <LoadingScreen />
        </div>
      </>
    );
  }

  return (
    <>
      <style>{css}</style>
      <div className="eca-root">
        <Header mode={mode} onSwitch={switchMode} />

        {mode === "voter" && voterStep === "login" && (
          <VoterLogin roster={roster} onVerified={(emp, record) => { setVerifiedVoter(emp); setExistingRecord(record); setVoterStep("voting"); }} />
        )}
        {mode === "voter" && voterStep === "voting" && verifiedVoter && (
          <VotingForm voter={verifiedVoter} categories={categories} existingRecord={existingRecord} onSubmitted={() => setVoterStep("done")} />
        )}
        {mode === "voter" && voterStep === "done" && <VoterDone />}

        {mode === "admin" && !adminAuthed && (
          <AdminLogin storedPasscode={passcode} onSetPasscode={handleSetPasscode} onLoggedIn={() => setAdminAuthed(true)} />
        )}
        {mode === "admin" && adminAuthed && (
          <AdminPanel
            roster={roster}
            setRoster={setRoster}
            categories={categories}
            setCategories={setCategories}
            storedPasscode={passcode}
            onSetPasscode={handleSetPasscode}
          />
        )}
      </div>
    </>
  );
}


const rootEl = document.getElementById("root");
const root = ReactDOM.createRoot(rootEl);
root.render(<App />);
