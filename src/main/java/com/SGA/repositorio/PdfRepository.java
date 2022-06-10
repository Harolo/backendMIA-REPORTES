package com.SGA.repositorio;

import java.util.List;

import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.CrudRepository;

import com.SGA.entidades.Estudiante;

public interface PdfRepository extends CrudRepository<Estudiante, Long>{

	@Query(value ="select id, nombre, apellidos, direccion, ciudad, edad from estudiante", nativeQuery = true)
	public List<Estudiante> findNamedeleteMailDireccion();
}
