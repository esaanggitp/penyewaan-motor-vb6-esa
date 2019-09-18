-- phpMyAdmin SQL Dump
-- version 4.5.1
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: 08 Jun 2017 pada 05.41
-- Versi Server: 10.1.9-MariaDB
-- PHP Version: 5.6.15

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `rental`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `detailsewa`
--

CREATE TABLE `detailsewa` (
  `id_sewa` varchar(20) NOT NULL,
  `jumlah_hari` int(11) NOT NULL,
  `subtotal` double NOT NULL,
  `m` int(5) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `detailsewa`
--

INSERT INTO `detailsewa` (`id_sewa`, `jumlah_hari`, `subtotal`, `m`) VALUES
('1705001', 1, 20000, 12),
('1705002', 3, 60000, 12),
('1705003', 1, 20000, 12),
('1705004', 2, 40000, 12),
('1705005', 1, 20000, 12),
('1705006', 2, 40000, 12),
('1705007', 3, 60000, 12),
('1705008', 1, 20000, 12),
('1705009', 2, 40000, 12),
('1706001', 2, 60000, 13),
('1706002', 2, 60000, 13);

-- --------------------------------------------------------

--
-- Struktur dari tabel `karyawan`
--

CREATE TABLE `karyawan` (
  `id_karyawan` int(10) NOT NULL,
  `password` varchar(10) NOT NULL,
  `nama_karyawan` varchar(30) NOT NULL,
  `alamat_karyawan` varchar(30) NOT NULL,
  `jenis_kelamin_karyawan` varchar(10) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `karyawan`
--

INSERT INTO `karyawan` (`id_karyawan`, `password`, `nama_karyawan`, `alamat_karyawan`, `jenis_kelamin_karyawan`) VALUES
(12151568, '12345678', 'esa anggit pangestu', 'banjarnegara', 'laki-laki');

-- --------------------------------------------------------

--
-- Struktur dari tabel `motor`
--

CREATE TABLE `motor` (
  `id_motor` int(5) NOT NULL,
  `no_plat` varchar(10) NOT NULL,
  `jenis` varchar(10) NOT NULL,
  `merk` varchar(10) NOT NULL,
  `thn_buat` int(4) NOT NULL,
  `warna` varchar(20) NOT NULL,
  `harga_motor` int(20) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `motor`
--

INSERT INTO `motor` (`id_motor`, `no_plat`, `jenis`, `merk`, `thn_buat`, `warna`, `harga_motor`) VALUES
(12, 'R-123KL', 'MTR_BEBEK', 'MIO', 2010, 'hitam', 20000),
(13, 'R-1234-JM', 'MTR_BEBEK', 'REVO', 2009, 'hitam', 30000);

-- --------------------------------------------------------

--
-- Struktur dari tabel `pelanggan`
--

CREATE TABLE `pelanggan` (
  `id_pelanggan` int(5) NOT NULL,
  `nama_pelanggan` varchar(30) NOT NULL,
  `alamat_pelanggan` varchar(30) NOT NULL,
  `jenis_kelamin_pelanggan` varchar(10) NOT NULL,
  `no_tlp_pelanggan` varchar(15) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `pelanggan`
--

INSERT INTO `pelanggan` (`id_pelanggan`, `nama_pelanggan`, `alamat_pelanggan`, `jenis_kelamin_pelanggan`, `no_tlp_pelanggan`) VALUES
(888, 'ANGELA', 'BANJARNEGARA', 'PEREMPUAN', '098878888'),
(999, 'PUTRI', 'BANJAR', 'PEREMPUAN', '00989999999');

-- --------------------------------------------------------

--
-- Struktur dari tabel `sewa`
--

CREATE TABLE `sewa` (
  `id_sewa` varchar(20) NOT NULL,
  `id_motor` int(5) NOT NULL,
  `tgl_pinjam` date NOT NULL,
  `total_bayar` int(10) NOT NULL,
  `id_karyawan` int(10) NOT NULL,
  `id_pelanggan` int(5) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `sewa`
--

INSERT INTO `sewa` (`id_sewa`, `id_motor`, `tgl_pinjam`, `total_bayar`, `id_karyawan`, `id_pelanggan`) VALUES
('1705001', 12, '2017-05-24', 20000, 0, 999),
('1705002', 12, '2017-05-24', 60000, 12151568, 999),
('1705003', 12, '2017-05-24', 20000, 12151568, 999),
('1705004', 12, '2017-05-24', 40000, 12151568, 888),
('1705005', 12, '2017-05-24', 20000, 12151568, 999),
('1705006', 12, '2017-05-29', 40000, 12151568, 999),
('1705007', 12, '2017-05-29', 60000, 12151568, 888),
('1705008', 12, '2017-05-29', 20000, 12151568, 999),
('1705009', 12, '2017-05-29', 40000, 12151568, 888),
('1706001', 13, '2017-06-04', 60000, 12151568, 999),
('1706002', 13, '2017-06-08', 60000, 12151568, 999);

--
-- Indexes for dumped tables
--

--
-- Indexes for table `detailsewa`
--
ALTER TABLE `detailsewa`
  ADD KEY `id_sewa` (`id_sewa`),
  ADD KEY `id_motor` (`jumlah_hari`),
  ADD KEY `id_sewa_2` (`id_sewa`);

--
-- Indexes for table `karyawan`
--
ALTER TABLE `karyawan`
  ADD PRIMARY KEY (`id_karyawan`);

--
-- Indexes for table `motor`
--
ALTER TABLE `motor`
  ADD PRIMARY KEY (`id_motor`,`no_plat`);

--
-- Indexes for table `pelanggan`
--
ALTER TABLE `pelanggan`
  ADD PRIMARY KEY (`id_pelanggan`);

--
-- Indexes for table `sewa`
--
ALTER TABLE `sewa`
  ADD PRIMARY KEY (`id_sewa`),
  ADD KEY `id_motor` (`id_motor`),
  ADD KEY `id_karyawan` (`id_karyawan`),
  ADD KEY `id_pelanggan` (`id_pelanggan`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
