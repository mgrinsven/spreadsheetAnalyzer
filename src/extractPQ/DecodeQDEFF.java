package extractPQ;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.ArrayList;

public class DecodeQDEFF {
	private static final boolean DEBUG=true;
	
	public static final int PACKAGEPARTS = 0;
	public static final int PERMISSIONS = 1;
	public static final int METADATA = 2;
	public static final int PERMISSIONSBINDING = 3;
	
	private static int FIELDS_LENGTH = 4;
	private ByteBuffer packageParts = null; 
	private ByteBuffer permissions = null; 
	private ByteBuffer metadata = null; 
	private ByteBuffer permissionsBinding = null;

	
	public void extractQDEFF(byte[] encoded, File output) {
		int version = ByteBuffer.wrap(encoded, 0, FIELDS_LENGTH).order(ByteOrder.LITTLE_ENDIAN).getInt();
	
		// Retrieve the size of the different parts of the PowerQuery data and calculate offsets
		int packagePartsOffset = FIELDS_LENGTH;
		int packagePartsLength = ByteBuffer.wrap(encoded, packagePartsOffset, FIELDS_LENGTH).order(ByteOrder.LITTLE_ENDIAN).getInt();
		
		int permissionsOffset = packagePartsOffset + FIELDS_LENGTH + packagePartsLength;
		int permissionsLength = ByteBuffer.wrap(encoded, permissionsOffset, FIELDS_LENGTH).order(ByteOrder.LITTLE_ENDIAN).getInt();
		
		int metadataOffset = permissionsOffset + FIELDS_LENGTH + permissionsLength;
		int metadataLength = ByteBuffer.wrap(encoded, metadataOffset, FIELDS_LENGTH).order(ByteOrder.LITTLE_ENDIAN).getInt();
		
		int permissionsBindingOffset = metadataOffset + FIELDS_LENGTH + metadataLength;
		int permissionsBindingLength = ByteBuffer.wrap(encoded, permissionsBindingOffset, FIELDS_LENGTH).order(ByteOrder.LITTLE_ENDIAN).getInt();
		
		// Allocate ByteBuffers of the different PowerQuery data parts
		packageParts = ByteBuffer.allocate(packagePartsLength);
		//permissions = ByteBuffer.allocate(permissionsLength);
		permissions = ByteBuffer.allocate(permissionsLength-3);
		//metadata = ByteBuffer.allocate(metadataLength);
		metadata = ByteBuffer.allocate(metadataLength-37);
		//permissionsBinding = ByteBuffer.allocate(permissionsBindingLength);
		permissionsBinding = ByteBuffer.allocate(permissionsBindingLength+26);
		
		// Fill allocated ByteBuffers with the different parts
		packageParts.put(encoded, packagePartsOffset + FIELDS_LENGTH, packagePartsLength);
		//permissions.put(encoded, permissionsOffset + FIELDS_LENGTH+3, permissionsLength);
		permissions.put(encoded, permissionsOffset + FIELDS_LENGTH+3, permissionsLength-3);
		metadata.put(encoded, metadataOffset + FIELDS_LENGTH+11, metadataLength-37);
		//permissionsBinding.put(encoded, permissionsBindingOffset + FIELDS_LENGTH, permissionsBindingLength);
		permissionsBinding.put(encoded, permissionsBindingOffset + FIELDS_LENGTH-26, permissionsBindingLength+26);
	}

	public boolean writeOutputFile(byte[] outputArray, File outputFile) {
		boolean result = true;
        try {
        	FileOutputStream fos = new FileOutputStream(outputFile);
        	fos.write(outputArray);
        	fos.close();
        } catch(IOException e) {
        	System.out.println(e.getMessage());
        	result=false;
        }
		return result;
	}
	
	public ByteBuffer getPackageParts() {
		return packageParts;
	}
	
	public ByteBuffer getPermissions() {
		return permissions;
	}
	
	public ByteBuffer getMetadata() {
		return metadata;
	}
	
	public ByteBuffer getPermissionsBinding() {
		return permissionsBinding;
	}
	
	public ArrayList<ByteBuffer> getPowerQueryParts() {
		ArrayList<ByteBuffer> pqByteBufferArray = new ArrayList<ByteBuffer>();
		pqByteBufferArray.add(PACKAGEPARTS, packageParts);
		pqByteBufferArray.add(PERMISSIONS, permissions);
		pqByteBufferArray.add(METADATA, metadata);
		pqByteBufferArray.add(PERMISSIONSBINDING, permissionsBinding);
		return pqByteBufferArray;
	}
}
